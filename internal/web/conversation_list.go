package web

import (
	"encoding/json"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"
)

var (
	testTitleRes = []*regexp.Regexp{
		regexp.MustCompile(`(?i)^say\s+["']?ok["']?\s+in\s+one\s+word\.?$`),
		regexp.MustCompile(`(?i)^只回复\s*ok\.?$`),
		regexp.MustCompile(`(?i)^hi\.?$`),
	}
	imageKindRe = regexp.MustCompile(`(?i)generate an image|gpt image|dall-?e|gpt-image|/v1/images|生图|画一张|画一个|image generation`)
	convCacheMu sync.Mutex
)

type conversationListCache struct {
	UpdatedAt string           `json:"updatedAt"`
	Items     []map[string]any `json:"items"`
}

func normalizeTestText(s string) string {
	s = strings.ReplaceAll(s, "\u201c", `"`)
	s = strings.ReplaceAll(s, "\u201d", `"`)
	s = strings.ReplaceAll(s, "\u2018", "'")
	s = strings.ReplaceAll(s, "\u2019", "'")
	s = strings.ReplaceAll(s, "[user]", " ")
	s = strings.ReplaceAll(s, "[assistant]", " ")
	s = strings.ReplaceAll(s, "[bot]", " ")
	s = strings.Join(strings.Fields(s), " ")
	return strings.TrimSpace(s)
}

func isTestConversationTitle(name string) bool {
	n := normalizeTestText(name)
	if n == "" {
		return false
	}
	for _, rx := range testTitleRes {
		if rx.MatchString(n) {
			return true
		}
	}
	return false
}

func isTestConversation(row map[string]any) bool {
	if row == nil {
		return false
	}
	if isTestConversationTitle(stringFromAny(row["chatName"]) + " " + stringFromAny(row["title"])) {
		return true
	}
	return false
}

func classifyConversationKind(row map[string]any) string {
	blob := strings.Join([]string{
		stringFromAny(row["chatName"]),
		stringFromAny(row["title"]),
		stringFromAny(row["endpoint"]),
		stringFromAny(row["model"]),
	}, " ")
	if imageKindRe.MatchString(blob) {
		return "image"
	}
	return "chat"
}

func stringFromAny(v any) string {
	s, _ := v.(string)
	return s
}

func isBlankAny(v any) bool {
	if v == nil {
		return true
	}
	s, ok := v.(string)
	return ok && s == ""
}

func conversationListCachePath() string {
	if dir := strings.TrimSpace(os.Getenv("M365_DATA_DIR")); dir != "" {
		return filepath.Join(dir, "conversation-list-cache.json")
	}
	if p := strings.TrimSpace(os.Getenv("M365_CONVERSATION_CACHE")); p != "" {
		return filepath.Join(filepath.Dir(p), "conversation-list-cache.json")
	}
	return filepath.Join("data", "conversation-list-cache.json")
}

func loadConversationListCache() conversationListCache {
	b, err := os.ReadFile(conversationListCachePath())
	if err != nil {
		return conversationListCache{}
	}
	var c conversationListCache
	if json.Unmarshal(b, &c) != nil || c.Items == nil {
		return conversationListCache{}
	}
	return c
}

func saveConversationListCache(items []map[string]any) conversationListCache {
	if items == nil {
		items = []map[string]any{}
	}
	c := conversationListCache{UpdatedAt: time.Now().UTC().Format(time.RFC3339), Items: items}
	b, err := json.MarshalIndent(c, "", "  ")
	if err != nil {
		return c
	}
	path := conversationListCachePath()
	_ = os.MkdirAll(filepath.Dir(path), 0o700)
	_ = os.WriteFile(path, b, 0o600)
	return c
}

func upsertConversationItems(oldItems, incoming []map[string]any) []map[string]any {
	byID := map[string]map[string]any{}
	for _, row := range oldItems {
		cid := strings.TrimSpace(stringFromAny(row["conversationId"]))
		if cid == "" || isTestConversation(row) {
			continue
		}
		byID[cid] = cloneRow(row)
	}
	for _, row := range incoming {
		if isTestConversation(row) {
			continue
		}
		cid := strings.TrimSpace(stringFromAny(row["conversationId"]))
		if cid == "" {
			continue
		}
		prev, ok := byID[cid]
		if !ok {
			byID[cid] = cloneRow(row)
			continue
		}
		if conversationTimestamp(row) >= conversationTimestamp(prev) {
			merged := cloneRow(prev)
			for k, v := range row {
				if isBlankAny(v) {
					continue
				}
				merged[k] = v
			}
			byID[cid] = merged
			continue
		}
		for _, k := range []string{"chatName", "accountEmail", "accountId", "messageCount"} {
			if isBlankAny(prev[k]) && !isBlankAny(row[k]) {
				prev[k] = row[k]
			}
		}
	}
	out := make([]map[string]any, 0, len(byID))
	for _, row := range byID {
		out = append(out, row)
	}
	return out
}

func cloneRow(row map[string]any) map[string]any {
	out := make(map[string]any, len(row))
	for k, v := range row {
		out[k] = v
	}
	return out
}

func (s *Server) listCloudConversationsAllAccounts() ([]map[string]any, error) {
	if s == nil || s.tokens == nil {
		return nil, nil
	}
	var (
		out    []map[string]any
		first  error
		listed bool
	)
	for _, acc := range s.tokens.List() {
		if strings.TrimSpace(acc.RefreshToken) == "" {
			continue
		}
		chats, err := s.cloudClientForAccount(acc).ListConversations()
		if err != nil {
			if first == nil {
				first = err
			}
			continue
		}
		listed = true
		for _, chat := range chats {
			if chat == nil {
				continue
			}
			cid := strings.TrimSpace(stringFromAny(chat["conversationId"]))
			if cid == "" {
				continue
			}
			chat["accountId"] = acc.ID
			chat["accountEmail"] = acc.Email
			chat["source"] = "m365-cloud"
			out = append(out, chat)
		}
	}
	if !listed && first != nil {
		return nil, first
	}
	return out, nil
}

func mergeConversationRows(rows ...[]map[string]any) map[string]map[string]any {
	out := map[string]map[string]any{}
	for _, group := range rows {
		for _, chat := range group {
			cid := strings.TrimSpace(stringFromAny(chat["conversationId"]))
			if cid == "" {
				continue
			}
			prev, ok := out[cid]
			if !ok {
				out[cid] = cloneRow(chat)
				continue
			}
			for k, v := range chat {
				if isBlankAny(v) {
					continue
				}
				if _, exists := prev[k]; !exists || isBlankAny(prev[k]) {
					prev[k] = v
				}
			}
		}
	}
	return out
}

func conversationRowsFromMerge(merged map[string]map[string]any) []map[string]any {
	out := make([]map[string]any, 0, len(merged))
	for _, row := range merged {
		if isTestConversation(row) {
			continue
		}
		out = append(out, row)
	}
	sort.Slice(out, func(i, j int) bool {
		return conversationTimestamp(out[i]) > conversationTimestamp(out[j])
	})
	return out
}

func (s *Server) collectGatewayConversationRows() []map[string]any {
	if s == nil || s.sessionResolver == nil {
		return nil
	}
	var out []map[string]any
	for _, session := range s.sessionResolver.ListSessions() {
		row := map[string]any{
			"conversationId":   session.ConversationID,
			"sessionId":        session.SessionID,
			"accountId":        session.AccountID,
			"createTimeUtc":    session.CreatedAt.UnixMilli(),
			"updateTimeUtc":    session.LastUsedAt.UnixMilli(),
			"messageCount":     len(session.ContextHistory),
			"historyAvailable": len(session.ContextHistory) > 0,
			"source":           "gateway",
			"chatName":         conversationTitle(session.ContextHistory),
		}
		if s.tokens != nil {
			if account, found := s.tokens.Get(session.AccountID); found {
				row["accountEmail"] = account.Email
			}
		}
		out = append(out, row)
	}
	return out
}

func paginateConversationRows(items []map[string]any, page, pageSize int) ([]map[string]any, int, int, int) {
	total := len(items)
	if page <= 0 {
		return items, 0, pageSize, total
	}
	if pageSize <= 0 {
		pageSize = 20
	}
	if pageSize > 100 {
		pageSize = 100
	}
	start := (page - 1) * pageSize
	if start >= total {
		return []map[string]any{}, page, pageSize, total
	}
	end := start + pageSize
	if end > total {
		end = total
	}
	return items[start:end], page, pageSize, total
}

func atoiDefault(s string, fallback int) int {
	n, err := strconv.Atoi(strings.TrimSpace(s))
	if err != nil {
		return fallback
	}
	return n
}

func dropConversationFromCache(id string) {
	id = strings.TrimSpace(id)
	if id == "" {
		return
	}
	convCacheMu.Lock()
	defer convCacheMu.Unlock()
	c := loadConversationListCache()
	out := make([]map[string]any, 0, len(c.Items))
	for _, row := range c.Items {
		if strings.TrimSpace(stringFromAny(row["conversationId"])) != id {
			out = append(out, row)
		}
	}
	saveConversationListCache(out)
}
