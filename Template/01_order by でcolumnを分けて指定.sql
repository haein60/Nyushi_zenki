/********************************************************/
/* order by でcolumnを分けて指定する例                  */
/* 2023.02.28 st jhi                                    */
/********************************************************/


select '1受験番号順',1 UNION SELECT '2出身高校順',2 UNION SELECT '3アイウエオ順',3
select '受験番号順','fp.iJukenNumber' UNION SELECT 'アイウエオ順','fp.vKanaName'


--for test変数
declare @vOrderField int;
set     @vOrderField=1;

select
    ROW_NUMBER() Over( ORDER BY CASE @vOrderField WHEN 1 THEN cast(dbo.usfMakeDispJukenNumber(iJukenNumber) as varchar(4)) WHEN 2 THEN vKanaName END)
   ,dbo.usfMakeDispJukenNumber(iJukenNumber) --受験NO
   ,vExamineeName
   ,vKanaName
-- ,vUpMark01 --性別
   ,CASE iSex WHEN 0 THEN 'M' WHEN 1 THEN 'F' ELSE '?' END iSex
-- ,vUpMark02 --年齢
   ,iAge
   ,vUpMark04 --出身県
   ,vUpMark05 --区分私立 or 公立
   ,SUBSTRING(vUpMark06,1,8) --高校情報
   ,vUpMark07 --評定値
   ,vUpMark08 --成績概評
from
    tbSTETempForPrint
where
        iReportNo='400200'
    and iNendo= 2023
    and iExamineeSTatus in(2,6)
    and iRejectFlag=0
   
ORDER BY
     --caseで指定するcolumnのtypeが同じではないとエラー発生 (eg)列1:int 列2:varcharの場合、エラーになる
     CASE @vOrderField 
         WHEN 1 THEN cast(dbo.usfMakeDispJukenNumber(iJukenNumber) as varchar(4)) --varchar
         WHEN 2 THEN vKanaName                                                    --varchar
     END
