-----------------------------------------------
exec uspSTESeisekiIchiran 400100,2023     

select top 500
    ROW_NUMBER() Over( ORDER BY fp.iJukenNumber)
   ,dbo.usfMakeDispJukenNumber(fp.iJukenNumber) --��NO
   ,fp.vExamineeName
   ,fp.vKanaName
-- ,fp.vUpMark01 --����
   ,CASE iSex WHEN 0 THEN 'M' WHEN 1 THEN 'F' ELSE '?' END iSex
-- ,fp.vUpMark02 --�N��
   ,fp.iAge
   ,fp.vUpMark04 --�o�g��
   ,fp.vUpMark05 --�敪���� or ����
   ,SUBSTRING(fp.vUpMark06,1,8) --���Z���
   ,fp.vUpMark07 --�]��l
   ,fp.vUpMark08 --���ъT�]
   ,fp.vUpMark09 --���i���
   ,fp.vUpMark10 --���t
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
   ,dbo.usfMakeDispJukenNumber(iJukenNumber) --��NO
   ,vExamineeName
   ,vKanaName
-- ,vUpMark01 --����
   ,CASE iSex WHEN 0 THEN 'M' WHEN 1 THEN 'F' ELSE '?' END iSex
-- ,vUpMark02 --�N��
   ,iAge
   ,vUpMark04 --�o�g��
   ,vUpMark05 --�敪���� or ����
   ,SUBSTRING(vUpMark06,1,8) --���Z���
   ,vUpMark07 --�]��l
   ,vUpMark08 --���ъT�]
   ,vUpMark09 --���i���
   ,vUpMark10 --���t
   ,vUpMark11 --���i���
   ,vUpMark12 --���t
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
/* SELECT DISTINCT ���w�肳��Ă���ꍇ�A�I�����X�g�� ORDER BY ���ڂ��K�v�ł��B */
/* �G���[�����������ꍇ�A�\������񖼂�K��1�w�肵�Ȃ���΂Ȃ�܂���          */
/* ���邢�� oder by����w�肵�Ȃ�(sub query�ɂ���)                              */
/* 2023.02.28 st jhi                                                            */ 
/********************************************************************************/

--delete from tbSTETempForPrint where  iReportNo='400700' and  iNendo= 2023 
exec uspSTESeisekiIchiran 400700,2023 

--for uspSTESeisekiIchiran exec�m�F
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

--distict�� ROW_NUMBER()...��(as sno)�̂�����distict�������Ȃ��̂�subquery���ɂ����B
(
 select distinct
    dbo.usfMakeDispJukenNumber(iJukenNumber) as vJukenNumber--��NO
   ,vExamineeName
   ,vKanaName
--,vUpMark01 --����
   ,CASE iSex WHEN 0 THEN 'M' WHEN 1 THEN 'F' ELSE '?' END iSex
--,vUpMark02 --�N��
   ,iAge
   ,vUpMark04 --�o�g��
   ,vUpMark05 --�敪���� or ����
   ,SUBSTRING(vUpMark06,1,8) as vUpMark06 --���Z���
   ,vUpMark07 --�]��l
   ,vUpMark08 --���ъT�]
   ,vUpMark09 --���i���
   ,vUpMark10 --���t
   ,vUpMark11 --���i���
   ,vUpMark12 --���t
from
    tbSTETempForPrint
where
        iReportNo='400700'
    and iNendo= 2023
    and iExamineeStatus=6
) bb

ORDER BY
    bb.vJukenNumber  
    
  