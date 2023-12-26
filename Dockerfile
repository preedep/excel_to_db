FROM alpine:3.19
RUN mkdir /app
WORKDIR /app
ADD ./target/x86_64-unknown-linux-musl/release/excel_to_db /app/excel_to_db
ENTRYPOINT ["/app/excel_to_db"]

