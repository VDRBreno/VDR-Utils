# Jaro Winkler - PostgreSQL

Função baseada em jaro winkler para analisar semelhanças entre duas strings. Escrito em PL/pgSQL, não é necessário nada externo.

### Como usar

Criar um script ex: `init_db.sql` contendo o conteúdo de `script.sql` e colocar no diretório de inicialização do Postgres, caso linux, é em `/docker-entrypoint-initdb.d`.

```sql
  SELECT * FROM table WHERE jaro_winkler(column, 'text test') > 0.5;

  WITH table_score AS (
    SELECT *, jaro_winkler(column, 'text test') > 0.5 as score FROM table
  ) SELECT * FROM table_score
    WHERE score > 0.5;
```