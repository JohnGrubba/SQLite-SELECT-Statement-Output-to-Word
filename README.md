# SQLITE Query Output To Word

- Requirements

You need a `.sqlite` Database and a `.sql` file with SQLite Querys in it. Every Query needs to have a Comment before it, for the program to work.

- Example for a `.sql` File

```sql
-- Query 1
SELECT * FROM fortnite_accs;
-- Query 2
SELECT your_penis FROM short_penises;
```
- Usage
```
python sql2word.py db.sqlite querys.sql output.docx
```
## Features
- Execute SQLite Querys and dump output into a .docx file

## Known Bugs
- Only one Command per Query