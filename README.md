# Excel2DB
This project developed with Rust programming language. for load data in my excel to database and support to use sql statements to query data in excel.


## Proof of concept 
- Read my excel (**demo.xlsx**) and save to database (SQLite memory)
- Show SQL prompt after save to database
- Can export result table to CSV file

### How to run 
```bash
docker run --rm -e RUST_LOG=info -v $(pwd):/data/ -it --entrypoint=/app/excel_to_db nickmsft/excel2db:latest -f /data/demo.xlsx -s Sheet1

```
### How to export result table to CSV file
```sql
[SQL] >> select * from excel_rows; |out=/data/test.csv

```




