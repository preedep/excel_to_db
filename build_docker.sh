echo "Building linux binary..."
./build_linux.sh
echo "Building docker image..."
docker rm -f excel2db:latest
docker build -t excel2db:latest .
