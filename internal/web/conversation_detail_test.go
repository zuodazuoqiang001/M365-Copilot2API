package web

import (
	"encoding/json"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"m365-copilot2api/internal/auth"
)

func TestConversationListAndDetailUseCompleteLocalHistory(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("M365_DATA_DIR", dir)
	t.Setenv("M365_SESSION_CACHE", filepath.Join(dir, "sessions.json"))
	t.Setenv("M365_CONVERSATION_CACHE", filepath.Join(dir, "conversations.json"))
	store, err := auth.OpenStore(filepath.Join(dir, "accounts.json"))
	if err != nil {
		t.Fatal(err)
	}
	s := &Server{tokens: store, sessionResolver: openSessionResolver()}

	oldCloudClient := m365CloudClient
	m365CloudClient = nil
	defer func() { m365CloudClient = oldCloudClient }()

	req := httptest.NewRequest(http.MethodPost, "/v1/chat/completions", nil)
	req.Header.Set(sessionHeaderName, "session-detail")
	body := &oaiReq{Messages: []oaiMsg{
		{Role: "user", Content: "show the complete answer"},
		{Role: "assistant", Content: "complete body", ReasoningContent: "complete reasoning"},
	}}
	s.sessionResolver.Bind("", "conversation-detail", "account-a", body, "", req)

	listRecorder := httptest.NewRecorder()
	s.handleM365Conversations(listRecorder, httptest.NewRequest(http.MethodGet, "/api/m365/conversations", nil))
	if listRecorder.Code != http.StatusOK {
		t.Fatalf("list status=%d body=%s", listRecorder.Code, listRecorder.Body.String())
	}
	var list struct {
		Count int              `json:"count"`
		Data  []map[string]any `json:"data"`
	}
	if err := json.Unmarshal(listRecorder.Body.Bytes(), &list); err != nil {
		t.Fatal(err)
	}
	if list.Count != 1 || list.Data[0]["messageCount"] != float64(2) {
		t.Fatalf("list response=%s", listRecorder.Body.String())
	}

	detailRecorder := httptest.NewRecorder()
	s.handleM365ConversationDetail(detailRecorder, httptest.NewRequest(http.MethodGet, "/api/m365/conversations/detail?id=conversation-detail", nil))
	if detailRecorder.Code != http.StatusOK {
		t.Fatalf("detail status=%d body=%s", detailRecorder.Code, detailRecorder.Body.String())
	}
	var detail struct {
		ConversationID string   `json:"conversationId"`
		Messages       []oaiMsg `json:"messages"`
	}
	if err := json.Unmarshal(detailRecorder.Body.Bytes(), &detail); err != nil {
		t.Fatal(err)
	}
	if detail.ConversationID != "conversation-detail" || len(detail.Messages) != 2 {
		t.Fatalf("detail response=%s", detailRecorder.Body.String())
	}
	if detail.Messages[1].ReasoningContent != "complete reasoning" || contentToString(detail.Messages[1].Content) != "complete body" {
		t.Fatalf("assistant message=%#v", detail.Messages[1])
	}
}

func TestConversationDetailPageContainsCompleteViews(t *testing.T) {
	body, err := os.ReadFile("../../web/conversation.html")
	if err != nil {
		t.Fatal(err)
	}
	page := string(body)
	for _, needle := range []string{
		`id="conversationView"`,
		`id="jsonView"`,
		"reasoning_content",
		"tool_calls",
		"/api/m365/conversations/detail?id=",
	} {
		if !strings.Contains(page, needle) {
			t.Fatalf("conversation page missing %q", needle)
		}
	}
}

func TestConversationTimestampPrefersUpdateTime(t *testing.T) {
	created := time.Now().Add(-time.Hour).UnixMilli()
	updated := time.Now().UnixMilli()
	if got := conversationTimestamp(map[string]any{"createTimeUtc": created, "updateTimeUtc": updated}); got != updated {
		t.Fatalf("timestamp=%d want %d", got, updated)
	}
}

func TestParseCloudConversationDetail(t *testing.T) {
	raw := []byte(`{"store":{"chatName":"store-name","rawConversationResponse":{"conversationId":"conv-cloud","chatName":"cloud-name","createTimeUtc":1,"updateTimeUtc":2,"messages":[{"author":"user","text":"hello"},{"author":"bot","text":"world"},{"author":"bot","text":"   "}]}}}`)
	got, err := parseCloudConversationDetail("fallback", raw)
	if err != nil {
		t.Fatal(err)
	}
	if got.ConversationID != "conv-cloud" || got.ChatName != "cloud-name" || len(got.Messages) != 2 {
		t.Fatalf("got=%+v", got)
	}
	if got.Messages[0].Role != "user" || contentToString(got.Messages[0].Content) != "hello" {
		t.Fatalf("user=%#v", got.Messages[0])
	}
	if got.Messages[1].Role != "assistant" || contentToString(got.Messages[1].Content) != "world" {
		t.Fatalf("assistant=%#v", got.Messages[1])
	}
}

func TestConversationDetailMissingWithoutCloudIsNotFound(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("M365_DATA_DIR", dir)
	t.Setenv("M365_SESSION_CACHE", filepath.Join(dir, "sessions.json"))
	t.Setenv("M365_CONVERSATION_CACHE", filepath.Join(dir, "conversations.json"))
	store, err := auth.OpenStore(filepath.Join(dir, "accounts.json"))
	if err != nil {
		t.Fatal(err)
	}
	s := &Server{tokens: store, sessionResolver: openSessionResolver()}
	oldCloudClient := m365CloudClient
	m365CloudClient = nil
	defer func() { m365CloudClient = oldCloudClient }()
	rec := httptest.NewRecorder()
	s.handleM365ConversationDetail(rec, httptest.NewRequest(http.MethodGet, "/api/m365/conversations/detail?id=missing", nil))
	if rec.Code != http.StatusNotFound {
		t.Fatalf("status=%d body=%s", rec.Code, rec.Body.String())
	}
}
