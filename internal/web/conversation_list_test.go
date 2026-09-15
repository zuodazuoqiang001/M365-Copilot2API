package web

import (
	"encoding/json"
	"net/http"
	"net/http/httptest"
	"path/filepath"
	"testing"

	"m365-copilot2api/internal/auth"
)

func TestIsTestConversationTitle(t *testing.T) {
	yes := []string{
		`Say "OK" in one word.`,
		"Say 'ok' in one word",
		"只回复ok",
		"只回复 ok.",
		"hi",
		"Hi.",
		`Say “OK” in one word.`,
	}
	for _, s := range yes {
		if !isTestConversationTitle(s) {
			t.Fatalf("expected test title %q", s)
		}
	}
	no := []string{"weekly report", "hello there", "生成一张猫的图片", ""}
	for _, s := range no {
		if isTestConversationTitle(s) {
			t.Fatalf("did not expect test title %q", s)
		}
	}
}

func TestClassifyConversationKind(t *testing.T) {
	if got := classifyConversationKind(map[string]any{"chatName": "Generate an image of a cat"}); got != "image" {
		t.Fatalf("kind=%s", got)
	}
	if got := classifyConversationKind(map[string]any{"endpoint": "/v1/images/generations"}); got != "image" {
		t.Fatalf("kind=%s", got)
	}
	if got := classifyConversationKind(map[string]any{"chatName": "weekly report"}); got != "chat" {
		t.Fatalf("kind=%s", got)
	}
}

func TestPaginateConversationRows(t *testing.T) {
	items := []map[string]any{{"id": "a"}, {"id": "b"}, {"id": "c"}}
	got, page, size, total := paginateConversationRows(items, 0, 20)
	if page != 0 || total != 3 || len(got) != 3 {
		t.Fatalf("unpaged page=%d size=%d total=%d n=%d", page, size, total, len(got))
	}
	got, page, size, total = paginateConversationRows(items, 2, 2)
	if page != 2 || size != 2 || total != 3 || len(got) != 1 {
		t.Fatalf("page=%d size=%d total=%d n=%d", page, size, total, len(got))
	}
}

func TestConversationListUsesCacheTabsAndFiltersTests(t *testing.T) {
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

	saveConversationListCache([]map[string]any{
		{"conversationId": "c-test", "chatName": `Say "OK" in one word.`, "updateTimeUtc": int64(3)},
		{"conversationId": "c-image", "chatName": "Generate an image of a cat", "updateTimeUtc": int64(2)},
		{"conversationId": "c-chat", "chatName": "weekly report", "updateTimeUtc": int64(1)},
	})

	chatRec := httptest.NewRecorder()
	s.handleM365Conversations(chatRec, httptest.NewRequest(http.MethodGet, "/api/m365/conversations?kind=chat&page=1&pageSize=20", nil))
	if chatRec.Code != http.StatusOK {
		t.Fatalf("chat status=%d body=%s", chatRec.Code, chatRec.Body.String())
	}
	var chat struct {
		Count  int              `json:"count"`
		Total  int              `json:"total"`
		Kind   string           `json:"kind"`
		Source string           `json:"source"`
		Totals map[string]int   `json:"totals"`
		Data   []map[string]any `json:"data"`
	}
	if err := json.Unmarshal(chatRec.Body.Bytes(), &chat); err != nil {
		t.Fatal(err)
	}
	if chat.Kind != "chat" || chat.Source != "cache" || chat.Count != 1 || chat.Total != 1 {
		t.Fatalf("chat list=%s", chatRec.Body.String())
	}
	if chat.Totals["chat"] != 1 || chat.Totals["image"] != 1 || chat.Totals["all"] != 2 {
		t.Fatalf("totals=%v body=%s", chat.Totals, chatRec.Body.String())
	}
	if chat.Data[0]["conversationId"] != "c-chat" {
		t.Fatalf("chat row=%v", chat.Data[0])
	}

	imageRec := httptest.NewRecorder()
	s.handleM365Conversations(imageRec, httptest.NewRequest(http.MethodGet, "/api/m365/conversations?kind=image&page=1&pageSize=20", nil))
	if imageRec.Code != http.StatusOK {
		t.Fatalf("image status=%d body=%s", imageRec.Code, imageRec.Body.String())
	}
	var image struct {
		Count int              `json:"count"`
		Data  []map[string]any `json:"data"`
	}
	if err := json.Unmarshal(imageRec.Body.Bytes(), &image); err != nil {
		t.Fatal(err)
	}
	if image.Count != 1 || image.Data[0]["conversationId"] != "c-image" {
		t.Fatalf("image list=%s", imageRec.Body.String())
	}
}

func TestDefaultTimeoutsAreFiveMinutes(t *testing.T) {
	t.Setenv("M365_CHAT_TIMEOUT_SECONDS", "")
	t.Setenv("M365_IMAGE_TIMEOUT_SECONDS", "")
	v := defaultRuntimeSettings()
	if v.ChatTimeoutSeconds != 300 || v.ImageTimeoutSeconds != 300 {
		t.Fatalf("timeouts chat=%d image=%d", v.ChatTimeoutSeconds, v.ImageTimeoutSeconds)
	}
}
