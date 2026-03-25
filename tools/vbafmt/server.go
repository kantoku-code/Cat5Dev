package main

import (
	"context"
	"encoding/json"
	"fmt"
	"net"
	"net/http"
	"os"
)

type formatRequest struct {
	Code    string  `json:"code"`
	Options Options `json:"options"`
}

type formatResponse struct {
	Result string `json:"result"`
	Error  string `json:"error"`
}

func startServer(port int) {
	mux := http.NewServeMux()
	mux.HandleFunc("/health", handleHealth)
	mux.HandleFunc("/format", handleFormat)

	listener, err := net.Listen("tcp", fmt.Sprintf("127.0.0.1:%d", port))
	if err != nil {
		fmt.Fprintf(os.Stderr, "vbafmt server: listen error: %v\n", err)
		os.Exit(1)
	}

	actualPort := listener.Addr().(*net.TCPAddr).Port
	// TypeScript 側がポート番号を読み取るために stdout に出力
	fmt.Printf("PORT=%d\n", actualPort)

	srv := &http.Server{Handler: mux}

	// SIGTERM / SIGINT でグレースフルシャットダウン
	go func() {
		waitSignal()
		srv.Shutdown(context.Background()) //nolint
	}()

	if err := srv.Serve(listener); err != nil && err != http.ErrServerClosed {
		fmt.Fprintf(os.Stderr, "vbafmt server: serve error: %v\n", err)
		os.Exit(1)
	}
}

func handleHealth(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "application/json")
	w.Write([]byte(`{"status":"ok"}`)) //nolint
}

func handleFormat(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
		return
	}

	var req formatRequest
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		w.Header().Set("Content-Type", "application/json")
		w.WriteHeader(http.StatusBadRequest)
		json.NewEncoder(w).Encode(formatResponse{Error: "invalid JSON: " + err.Error()}) //nolint
		return
	}

	result, fmtErr := safeFormat(req.Code, req.Options)

	w.Header().Set("Content-Type", "application/json")
	resp := formatResponse{Result: result}
	if fmtErr != "" {
		w.WriteHeader(http.StatusInternalServerError)
		resp.Error = fmtErr
	}
	json.NewEncoder(w).Encode(resp) //nolint
}

// safeFormat はパニックを recover して文字列エラーとして返す
func safeFormat(code string, opts Options) (result string, errMsg string) {
	defer func() {
		if r := recover(); r != nil {
			errMsg = fmt.Sprintf("format panic: %v", r)
		}
	}()
	return Format(code, opts), ""
}
