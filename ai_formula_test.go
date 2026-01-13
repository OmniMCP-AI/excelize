package excelize

import (
	"container/list"
	"io"
	"net/http"
	"strings"
	"testing"
)

type roundTripFunc func(*http.Request) (*http.Response, error)

func (fn roundTripFunc) RoundTrip(req *http.Request) (*http.Response, error) {
	return fn(req)
}

func TestAIInvalidArgs(t *testing.T) {
	fn := &formulaFuncs{}
	args := list.New()
	args.PushBack(newStringFormulaArg("excel__tool"))
	args.PushBack(newStringFormulaArg(`{"bad json"`))

	result := fn.AI(args)
	if result.Type != ArgString || !strings.Contains(result.String, "ERROR") {
		t.Fatalf("expected JSON error, got %#v", result)
	}
}

func TestCallFastestAIAPISuccess(t *testing.T) {
	origTransport := http.DefaultTransport
	defer func() {
		http.DefaultTransport = origTransport
	}()

	http.DefaultTransport = roundTripFunc(func(req *http.Request) (*http.Response, error) {
		if req.URL.String() != "https://api.fastest.ai/v1/tool/function_call" {
			t.Fatalf("unexpected URL: %s", req.URL)
		}
		body := `{"success":true,"cell_value":"42"}`
		return &http.Response{
			StatusCode: http.StatusOK,
			Body:       io.NopCloser(strings.NewReader(body)),
			Header:     make(http.Header),
		}, nil
	})

	fn := &formulaFuncs{sheet: "SheetX", cell: "B2"}
	result, err := callFastestAIAPI(fn, "excel__read", map[string]interface{}{"uri": "https://example.com"})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if result != "42" {
		t.Fatalf("unexpected result %q", result)
	}
}

func TestCallFastestAIAPIError(t *testing.T) {
	origTransport := http.DefaultTransport
	defer func() {
		http.DefaultTransport = origTransport
	}()

	http.DefaultTransport = roundTripFunc(func(req *http.Request) (*http.Response, error) {
		body := `{"success":false,"error":"boom"}`
		return &http.Response{
			StatusCode: http.StatusOK,
			Body:       io.NopCloser(strings.NewReader(body)),
			Header:     make(http.Header),
		}, nil
	})

	fn := &formulaFuncs{sheet: "Sheet1"}
	if _, err := callFastestAIAPI(fn, "excel__tool", map[string]interface{}{}); err == nil {
		t.Fatalf("expected error")
	}
}
