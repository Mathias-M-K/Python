create table testdb.unionTable
SELECT team,W,L,T,League,'year','W-L%' from testDb.assignment_2_nfl
union all
select team,W,L,'N/A',League,'year','W-L%' from testDb.assignment_2_nba