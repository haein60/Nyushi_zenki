/********************************************************/
/* order by ��column�𕪂��Ďw�肷���                  */
/* 2023.02.28 st jhi                                    */
/********************************************************/


select '1�󌱔ԍ���',1 UNION SELECT '2�o�g���Z��',2 UNION SELECT '3�A�C�E�G�I��',3
select '�󌱔ԍ���','fp.iJukenNumber' UNION SELECT '�A�C�E�G�I��','fp.vKanaName'


--for test�ϐ�
declare @vOrderField int;
set     @vOrderField=1;

select
    ROW_NUMBER() Over( ORDER BY CASE @vOrderField WHEN 1 THEN cast(dbo.usfMakeDispJukenNumber(iJukenNumber) as varchar(4)) WHEN 2 THEN vKanaName END)
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
from
    tbSTETempForPrint
where
        iReportNo='400200'
    and iNendo= 2023
    and iExamineeSTatus in(2,6)
    and iRejectFlag=0
   
ORDER BY
     --case�Ŏw�肷��column��type�������ł͂Ȃ��ƃG���[���� (eg)��1:int ��2:varchar�̏ꍇ�A�G���[�ɂȂ�
     CASE @vOrderField 
         WHEN 1 THEN cast(dbo.usfMakeDispJukenNumber(iJukenNumber) as varchar(4)) --varchar
         WHEN 2 THEN vKanaName                                                    --varchar
     END
