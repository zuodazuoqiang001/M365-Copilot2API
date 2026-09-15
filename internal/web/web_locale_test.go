package web

import (
	"os"
	"strings"
	"testing"
)

func TestWebIndexDefaultsToChineseUntilLocaleIsSelected(t *testing.T) {
	body, err := os.ReadFile("../../web/index.html")
	if err != nil {
		t.Fatal(err)
	}
	page := string(body)
	for _, needle := range []string{
		"const localeSelectionKey='m365_locale_selected';",
		"function preferredLocale()",
		"return 'zh-CN';",
		"localStorage.setItem(localeSelectionKey,'1')",
	} {
		if !strings.Contains(page, needle) {
			t.Fatalf("web index missing Chinese default bootstrap %q", needle)
		}
	}
}

func TestWebIndexIncludesConversationTabsAndQuotaRefresh(t *testing.T) {
	body, err := os.ReadFile("../../web/index.html")
	if err != nil {
		t.Fatal(err)
	}
	page := string(body)
	for _, needle := range []string{
		`id="acctRefreshAllBtn"`,
		`id="convRefreshBtn"`,
		`id="convTabChat"`,
		`id="convTabImage"`,
		`id="convPager"`,
		`id="setImageTimeout"`,
		`function refreshAllQuotas()`,
		`kind:convKind||'chat'`,
	} {
		if !strings.Contains(page, needle) {
			t.Fatalf("web index missing conversation/quota UI %q", needle)
		}
	}
}

func TestWebIndexIncludesAccountMonitoringControls(t *testing.T) {
	body, err := os.ReadFile("../../web/index.html")
	if err != nil {
		t.Fatal(err)
	}
	page := string(body)
	for _, needle := range []string{
		`data-f="cooldown"`,
		`x.status==='cooldown'`,
		`/api/accounts/schedule`,
		`x.callCount||0`,
		`x.rateLimited`,
		`Limited after ${x.callCount||0} calls`,
	} {
		if !strings.Contains(page, needle) {
			t.Fatalf("web index missing cooldown control %q", needle)
		}
	}
}
