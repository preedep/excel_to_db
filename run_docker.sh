docker run --rm -e RUST_LOG=info -v $(pwd):/data/ -it --entrypoint=/app/excel_to_db nickmsft/excel2db:latest -f /data/demo.xlsx -s Sheet1
