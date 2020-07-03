mac:
	go build -o target/mac/日計プログラム ./*.go

windows:
	GOOS=windows GOARCH=amd64 go build -o target/windows-amd64/日計プログラム.exe ./*.go