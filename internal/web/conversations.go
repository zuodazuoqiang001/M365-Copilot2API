package web

import (
	"encoding/json"
	"log"
	"net/http"
	"os"
	"strings"
	"time"

	"m365-copilot2api/internal/auth"
)

func (s *Server) conversations(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodGet {
		writeOpenAIError(w, http.StatusMethodNotAllowed, "invalid_request_error", "method not allowed")
		return
	}
	jsonOut(w, map[string]any{"conversations": s.sessions.list()})
}

func (s *Server) deleteConversation(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		writeOpenAIError(w, http.StatusMethodNotAllowed, "invalid_request_error", "method not allowed")
		return
	}
	var body struct {
		ID string `json:"id"`
	}
	if json.NewDecoder(r.Body).Decode(&body) != nil || body.ID == "" {
		writeOpenAIError(w, http.StatusBadRequest, "invalid_request_error", "bad json")
		return
	}
	s.conversationManager.Delete(body.ID)
	s.sessions.delete(body.ID)
	jsonOut(w, map[string]string{"status": "deleted"})
}

func (s *Server) conversationCleanup(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		writeOpenAIError(w, http.StatusMethodNotAllowed, "invalid_request_error", "method not allowed")
		return
	}
	var body struct {
		Mode  string `json:"mode"`
		KeepN int    `json:"keep_n"`
	}
	if json.NewDecoder(r.Body).Decode(&body) == nil {
		if body.Mode != "" {
			s.conversationManager.SetMode(ConversationCleanupMode(body.Mode))
		}
	}
	cleaned := s.conversationManager.Cleanup()
	jsonOut(w, map[string]any{
		"status":    "cleaned",
		"mode":      string(s.conversationManager.Mode()),
		"deleted":   cleaned,
		"remaining": len(s.conversationManager.List()),
	})
}

// publicSessionID returns the identifier a client uses to refer to a binding:
// its own explicit X-M365-Session-Id when it set one, otherwise the internal
// session id. The tenant hash and stored conversation history are never
// exposed through the API.
func publicSessionID(sess sessionBinding) string {
	if sess.ExplicitID != "" {
		return sess.ExplicitID
	}
	return sess.SessionID
}

func (s *Server) handleSessions(w http.ResponseWriter, r *http.Request) {
	tenant := tenantFromRequest(r)
	switch r.Method {
	case http.MethodGet:
		sessions := s.sessionResolver.ListSessionsForTenant(tenant)
		data := make([]map[string]any, 0, len(sessions))
		for _, sess := range sessions {
			data = append(data, map[string]any{
				"id":              publicSessionID(sess),
				"conversation_id": sess.ConversationID,
				"created":         sess.CreatedAt.Unix(),
				"last_used":       sess.LastUsedAt.Unix(),
				"messages":        len(sess.ContextHistory),
			})
		}
		jsonOut(w, map[string]any{
			"object": "list",
			"data":   data,
		})
	case http.MethodPost:
		var body struct {
			SessionID string `json:"session_id"`
		}
		json.NewDecoder(r.Body).Decode(&body)
		sess, ok := s.sessionResolver.GetSession(tenant, body.SessionID)
		if !ok {
			jsonOut(w, map[string]any{
				"object":     "session",
				"id":         body.SessionID,
				"created":    time.Now().Unix(),
				"expires_in": 1800,
				"status":     "created",
			})
			return
		}
		jsonOut(w, map[string]any{
			"object":          "session",
			"id":              publicSessionID(sess),
			"conversation_id": sess.ConversationID,
			"created":         sess.CreatedAt.Unix(),
			"status":          "active",
		})
	default:
		writeOpenAIError(w, http.StatusMethodNotAllowed, "invalid_request_error", "method not allowed")
	}
}

func (s *Server) handleCacheStats(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodGet {
		writeOpenAIError(w, http.StatusMethodNotAllowed, "invalid_request_error", "method not allowed")
		return
	}
	stats := cacheStats.GetStats()
	jsonOut(w, map[string]any{
		"object":     "cache_stats",
		"stats":      stats,
		"conv_cache": s.convCache.Stats(),
	})
}

func (s *Server) handleCacheStatsReset(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		writeOpenAIError(w, http.StatusMethodNotAllowed, "invalid_request_error", "method not allowed")
		return
	}
	cacheStats.Reset()
	jsonOut(w, map[string]any{"status": "reset"})
}

func (s *Server) handleM365Conversations(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodGet {
		writeOpenAIError(w, http.StatusMethodNotAllowed, "invalid_request_error", "method not allowed")
		return
	}
	refresh := r.URL.Query().Get("refresh") == "1" || r.URL.Query().Get("refresh") == "true"
	kind := strings.ToLower(strings.TrimSpace(r.URL.Query().Get("kind")))
	if kind == "" {
		kind = "all"
	}
	page := atoiDefault(r.URL.Query().Get("page"), 0)
	pageSize := atoiDefault(r.URL.Query().Get("pageSize"), 20)

	local := s.collectGatewayConversationRows()
	convCacheMu.Lock()
	cache := loadConversationListCache()
	hasCache := len(cache.Items) > 0
	convCacheMu.Unlock()

	noAccounts := s.tokens == nil || len(s.tokens.List()) == 0
	if m365CloudClient == nil && len(local) == 0 && noAccounts && (refresh || !hasCache) {
		writeOpenAIError(w, http.StatusServiceUnavailable, "m365_not_configured", "M365 cloud client not configured. Please add an M365 account first via PKCE authorization.")
		return
	}

	needRefresh := refresh || !hasCache
	var cloud []map[string]any
	var cloudErr error
	source := "cache"
	if needRefresh {
		cloud, cloudErr = s.listCloudConversationsAllAccounts()
		if cloud == nil && m365CloudClient != nil {
			if chats, e2 := m365CloudClient.ListConversations(); e2 == nil {
				cloud = chats
			} else if cloudErr == nil {
				cloudErr = e2
			}
		}
		source = "refresh"
	}

	convCacheMu.Lock()
	cache = loadConversationListCache()
	items := cache.Items
	if needRefresh {
		items = upsertConversationItems(items, cloud)
	}
	items = conversationRowsFromMerge(mergeConversationRows(items, local))
	if needRefresh {
		cache = saveConversationListCache(items)
	} else {
		cache.Items = items
	}
	convCacheMu.Unlock()

	if cloudErr != nil && len(cache.Items) == 0 && len(local) == 0 {
		writeOpenAIError(w, http.StatusBadGateway, "m365_error", cloudErr.Error())
		return
	}

	chatN, imageN := 0, 0
	filtered := make([]map[string]any, 0, len(cache.Items))
	for _, row := range cache.Items {
		if isTestConversation(row) {
			continue
		}
		k := classifyConversationKind(row)
		row["kind"] = k
		if k == "image" {
			imageN++
		} else {
			chatN++
		}
		if kind == "all" || kind == k {
			filtered = append(filtered, row)
		}
	}
	pageRows, page, pageSize, total := paginateConversationRows(filtered, page, pageSize)
	response := map[string]any{
		"object":             "list",
		"data":               pageRows,
		"count":              len(pageRows),
		"total":              total,
		"page":               page,
		"pageSize":           pageSize,
		"source":             source,
		"refreshedAt":        cache.UpdatedAt,
		"kind":               kind,
		"totals":             map[string]int{"all": chatN + imageN, "chat": chatN, "image": imageN},
		"aggregatedAccounts": true,
	}
	if cloudErr != nil {
		response["warning"] = cloudErr.Error()
	}
	jsonOut(w, response)
}

func (s *Server) handleM365ConversationDetail(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodGet {
		writeOpenAIError(w, http.StatusMethodNotAllowed, "invalid_request_error", "method not allowed")
		return
	}
	conversationID := strings.TrimSpace(r.URL.Query().Get("id"))
	if conversationID == "" {
		writeOpenAIError(w, http.StatusBadRequest, "invalid_request_error", "conversation id is required")
		return
	}
	session, found := s.sessionResolver.GetConversation(conversationID)
	if found && len(session.ContextHistory) > 0 {
		writeLocalConversationDetail(w, s, session)
		return
	}
	if detail, acc, ok := s.lookupCloudConversationDetail(conversationID); ok {
		log.Printf("[m365-cloud] conversation detail fallback id=%s messages=%d", conversationID, len(detail.Messages))
		jsonOut(w, map[string]any{
			"object":         "conversation",
			"conversationId": detail.ConversationID,
			"sessionId":      "",
			"accountId":      acc.ID,
			"accountEmail":   acc.Email,
			"chatName":       detail.ChatName,
			"createdAt":      detail.CreatedAt,
			"updatedAt":      detail.UpdatedAt,
			"messageCount":   len(detail.Messages),
			"messages":       detail.Messages,
			"source":         "m365-cloud",
		})
		return
	}
	if found {
		writeLocalConversationDetail(w, s, session)
		return
	}
	writeOpenAIError(w, http.StatusNotFound, "conversation_not_found", "conversation history is not available")
}

func writeLocalConversationDetail(w http.ResponseWriter, s *Server, session sessionBinding) {
	accountEmail := ""
	if s != nil && s.tokens != nil {
		if account, ok := s.tokens.Get(session.AccountID); ok {
			accountEmail = account.Email
		}
	}
	jsonOut(w, map[string]any{
		"object":         "conversation",
		"conversationId": session.ConversationID,
		"sessionId":      session.SessionID,
		"accountId":      session.AccountID,
		"accountEmail":   accountEmail,
		"chatName":       conversationTitle(session.ContextHistory),
		"createdAt":      session.CreatedAt,
		"updatedAt":      session.LastUsedAt,
		"messageCount":   len(session.ContextHistory),
		"messages":       session.ContextHistory,
		"source":         "gateway",
	})
}

func (s *Server) cloudClientForAccount(acc auth.AccountToken) *M365CloudClient {
	clientID := strings.TrimSpace(os.Getenv("M365_CLIENT_ID"))
	if clientID == "" {
		clientID = acc.ClientID
	}
	if clientID == "" {
		clientID = auth.DefaultClientID
	}
	tid := acc.TID
	if tid == "" {
		tid = "common"
	}
	return NewM365CloudClient(clientID, tid, acc.RefreshToken)
}

func (s *Server) lookupCloudConversationDetail(conversationID string) (cloudConversationDetail, auth.AccountToken, bool) {
	if s == nil || s.tokens == nil {
		return cloudConversationDetail{}, auth.AccountToken{}, false
	}
	for _, acc := range s.tokens.List() {
		if strings.TrimSpace(acc.RefreshToken) == "" {
			continue
		}
		detail, err := s.cloudClientForAccount(acc).GetConversation(conversationID)
		if err != nil {
			continue
		}
		return detail, acc, true
	}
	return cloudConversationDetail{}, auth.AccountToken{}, false
}

func conversationTitle(messages []oaiMsg) string {
	for _, message := range messages {
		if message.Role != "user" {
			continue
		}
		text := strings.TrimSpace(contentToString(message.Content))
		text = strings.Join(strings.Fields(text), " ")
		if text == "" {
			continue
		}
		runes := []rune(text)
		if len(runes) > 120 {
			return string(runes[:120]) + "..."
		}
		return text
	}
	return "Untitled conversation"
}

func conversationTimestamp(row map[string]any) int64 {
	for _, key := range []string{"updateTimeUtc", "createTimeUtc"} {
		switch value := row[key].(type) {
		case float64:
			return int64(value)
		case int64:
			return value
		case int:
			return int64(value)
		}
	}
	return 0
}

func (s *Server) handleM365Delete(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		writeOpenAIError(w, http.StatusMethodNotAllowed, "invalid_request_error", "method not allowed")
		return
	}
	var body struct {
		ConversationID string `json:"conversation_id"`
	}
	if json.NewDecoder(r.Body).Decode(&body) != nil || body.ConversationID == "" {
		writeOpenAIError(w, http.StatusBadRequest, "invalid_request_error", "bad json")
		return
	}
	var delErr error
	if m365CloudClient != nil {
		delErr = m365CloudClient.DeleteConversation(body.ConversationID)
	}
	if s.tokens != nil {
		for _, acc := range s.tokens.List() {
			if strings.TrimSpace(acc.RefreshToken) == "" {
				continue
			}
			if err := s.cloudClientForAccount(acc).DeleteConversation(body.ConversationID); err != nil {
				if delErr == nil {
					delErr = err
				}
				continue
			}
			delErr = nil
			break
		}
	}
	if m365CloudClient == nil && (s.tokens == nil || len(s.tokens.List()) == 0) {
		writeOpenAIError(w, http.StatusServiceUnavailable, "m365_not_configured", "M365 cloud client not configured. Please add an M365 account first via PKCE authorization.")
		return
	}
	if delErr != nil && m365CloudClient == nil {
		writeOpenAIError(w, http.StatusBadGateway, "m365_error", delErr.Error())
		return
	}
	s.dropConversation(body.ConversationID)
	dropConversationFromCache(body.ConversationID)
	jsonOut(w, map[string]any{"status": "deleted", "conversation_id": body.ConversationID})
}

func (s *Server) handleM365Cleanup(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		writeOpenAIError(w, http.StatusMethodNotAllowed, "invalid_request_error", "method not allowed")
		return
	}
	if m365CloudClient == nil {
		writeOpenAIError(w, http.StatusServiceUnavailable, "m365_not_configured", "M365 cloud client not configured. Please add an M365 account first via PKCE authorization.")
		return
	}
	var body struct {
		MaxAgeHours int `json:"max_age_hours"`
		KeepN       int `json:"keep_n"`
	}
	json.NewDecoder(r.Body).Decode(&body)

	maxAge := time.Duration(body.MaxAgeHours) * time.Hour
	if maxAge <= 0 {
		maxAge = 24 * time.Hour
	}
	keepN := body.KeepN
	if keepN <= 0 {
		keepN = 5
	}

	deleted, err := m365CloudClient.CleanupOldConversations(maxAge, keepN)
	if err != nil {
		writeOpenAIError(w, http.StatusBadGateway, "m365_error", err.Error())
		return
	}
	jsonOut(w, map[string]any{"status": "cleaned", "deleted": deleted})
}

func (s *Server) handleSessionDelete(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodDelete {
		writeOpenAIError(w, http.StatusMethodNotAllowed, "invalid_request_error", "method not allowed")
		return
	}
	sessionID := strings.TrimPrefix(r.URL.Path, "/v1/sessions/")
	if sessionID == "" {
		writeOpenAIError(w, http.StatusBadRequest, "invalid_request_error", "session_id required")
		return
	}
	if s.sessionResolver.DeleteSession(tenantFromRequest(r), sessionID) {
		jsonOut(w, map[string]any{"status": "deleted", "session_id": sessionID})
	} else {
		writeOpenAIError(w, http.StatusNotFound, "not_found", "session not found")
	}
}

type conversationWhitelistRequest struct {
	ConversationID string `json:"conversation_id"`
	Add            bool   `json:"add"`
}

func (s *Server) conversationWhitelist(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		writeOpenAIError(w, http.StatusMethodNotAllowed, "invalid_request_error", "method not allowed")
		return
	}
	var body conversationWhitelistRequest
	if json.NewDecoder(r.Body).Decode(&body) != nil || body.ConversationID == "" {
		writeOpenAIError(w, http.StatusBadRequest, "invalid_request_error", "bad json")
		return
	}
	if body.Add {
		s.conversationManager.Whitelist(body.ConversationID)
	} else {
		s.conversationManager.Unwhitelist(body.ConversationID)
	}
	jsonOut(w, map[string]any{"status": "updated", "conversation_id": body.ConversationID, "whitelisted": body.Add})
}
