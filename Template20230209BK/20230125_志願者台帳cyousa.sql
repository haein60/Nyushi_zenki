--exec uspSTESeisekiIchiran2020	100100,2023

Select 
ROW_NUMBER() Over( ORDER BY vKanaName)
, dbo.usfMakeDispJukenNumber(iJukenNumber)
,vExamineeName
,vKanaName,vUpMark01
,vUpMark02,vUpMark04
,vUpMark05,vUpMark06
,vUpMark07,vUpMark08
FROM
    tbSTETempForPrint WITH(NOLOCK)
Where
        ireportNO='100100'
    AND iNendo=2023
    AND iJukenNumber Between '1' AND '2354'
 --order by
 --    vKanaName