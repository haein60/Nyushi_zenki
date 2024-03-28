-----------------------------------------------
exec uspSTESeisekiIchiran 400100,2023     

select top 500
    ROW_NUMBER() Over( ORDER BY fp.iJukenNumber)
   ,dbo.usfMakeDispJukenNumber(fp.iJukenNumber) --受験NO
   ,fp.vExamineeName
   ,fp.vKanaName
-- ,fp.vUpMark01 --性別
   ,CASE iSex WHEN 0 THEN 'M' WHEN 1 THEN 'F' ELSE '?' END iSex
-- ,fp.vUpMark02 --年齢
   ,fp.iAge
   ,fp.vUpMark04 --出身県
   ,fp.vUpMark05 --区分私立 or 公立
   ,SUBSTRING(fp.vUpMark06,1,8) --高校情報
   ,fp.vUpMark07 --評定値
   ,fp.vUpMark08 --成績概評
   ,fp.vUpMark09 --合格状態
   ,fp.vUpMark10 --日付
from
    tbSTETempForPrint fp
where
        fp.ireportNO='400100'
    and fp.iNendo= 2022
    and fp.iExamineeSTatus in(2,6)
    and fp.iRejectFlag=0
ORDER BY
   fp.iJukenNumber 
   
-----------------------------------------------
exec uspSTESeisekiIchiran 400600,2023     
select
    ROW_NUMBER() Over( ORDER BY iJukenNumber)
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
   ,vUpMark09 --合格状態
   ,vUpMark10 --日付
   ,vUpMark11 --合格状態
   ,vUpMark12 --日付
from
    tbSTETempForPrint
where
        iReportNo='400600'
    and iNendo= 2023
    and iExamineeSTatus in(2,6)
    and iRejectFlag=1
ORDER BY
    iJukenNumber  


/********************************************************************************/
/* SELECT DISTINCT が指定されている場合、選択リストに ORDER BY 項目が必要です。 */
/* エラーが発生した場合、表示する列名を必ず1個指定しなければなりません          */
/* あるいは oder by句を指定しない(sub queryにする)                              */
/* 2023.02.28 st jhi                                                            */ 
/********************************************************************************/

--delete from tbSTETempForPrint where  iReportNo='400700' and  iNendo= 2023 
exec uspSTESeisekiIchiran 400700,2023 

--for uspSTESeisekiIchiran exec確認
select distinct
    *
from
    tbSTETempForPrint 
where
        iReportNo='400700'
    and iNendo= 2023
    and iExamineeSTatus=6
    
    
 
select
    ROW_NUMBER() Over( ORDER BY bb.vJukenNumber) sno
   ,bb.*
from

--distictが ROW_NUMBER()...列(as sno)のせいでdistictが聞かないのでsubquery化にした。
(
 select distinct
    dbo.usfMakeDispJukenNumber(iJukenNumber) as vJukenNumber--受験NO
   ,vExamineeName
   ,vKanaName
--,vUpMark01 --性別
   ,CASE iSex WHEN 0 THEN 'M' WHEN 1 THEN 'F' ELSE '?' END iSex
--,vUpMark02 --年齢
   ,iAge
   ,vUpMark04 --出身県
   ,vUpMark05 --区分私立 or 公立
   ,SUBSTRING(vUpMark06,1,8) as vUpMark06 --高校情報
   ,vUpMark07 --評定値
   ,vUpMark08 --成績概評
   ,vUpMark09 --合格状態
   ,vUpMark10 --日付
   ,vUpMark11 --合格状態
   ,vUpMark12 --日付
from
    tbSTETempForPrint
where
        iReportNo='400700'
    and iNendo= 2023
    and iExamineeStatus=6
) bb

ORDER BY
    bb.vJukenNumber  
    
  