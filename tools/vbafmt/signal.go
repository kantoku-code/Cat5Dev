package main

import (
	"os"
	"os/signal"
	"syscall"
)

// waitSignal は SIGINT または SIGTERM を受信するまでブロックする
func waitSignal() {
	ch := make(chan os.Signal, 1)
	signal.Notify(ch, syscall.SIGINT, syscall.SIGTERM)
	<-ch
}
