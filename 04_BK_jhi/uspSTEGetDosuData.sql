if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[uspSTEGetDosuData]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[uspSTEGetDosuData]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE uspSTEGetDosuData
	@vNendo	varchar(4)
,	@vTotalCategoryID	varchar(4)
,	@vPrintStartScore		varchar(4)
,	@vPrintEndScore		varchar(4)
,	@vScore		varchar(4)
,	@vTargetExamineeCategoryID1	varchar(4)
,	@vTargetExamineeAdmission1	varchar(2)
,	@vTargetExamineeSex1	char(2)
,	@vTargetExamineeCategoryID2	varchar(4) = ''
,	@vTargetExamineeAdmission2	varchar(2) = ''
,	@vTargetExamineeSex2	char(2) = ''
,	@vTargetExamineeCategoryID3	varchar(4) = ''
,	@vTargetExamineeAdmission3	varchar(2) = ''
,	@vTargetExamineeSex3	char(2) = ''
 AS
declare
	@fStartScore	float
,	@iTotalCategoryID	int
,	@iPrintStartScore		int
,	@iPrintEndScore		int
,	@fScore		float
,	@iTargetExamineeCategoryID1 int
,	@iTargetExamineeAdmission1 int
,	@iTargetExamineeSex1 int
,	@iTargetExamineeCategoryID2	int
,	@iTargetExamineeAdmission2 int
,	@iTargetExamineeSex2 int
,	@iTargetExamineeCategoryID3	int
,	@iTargetExamineeAdmission3 int
,	@iTargetExamineeSex3 int
,	@fAvg1	float
,	@fAvg2	float
,	@fAvg3	float
,	@fSum1	float
,	@fSum2	float
,	@fSum3	float
,	@fMax1	float
,	@fMax2	float
,	@fMax3	float
,	@fMin1	float
,	@fMin2	float
,	@fMin3	float
,	@fSd1	float
,	@fSd2	float
,	@fSd3	float
,	@lCnt1	int
,	@lCnt2	int
,	@lCnt3	int

begin
    set @iTotalCategoryID = convert( int , @vTotalCategoryID )
    set @iPrintStartScore = convert( int , @vPrintStartScore )
    set @iPrintEndScore = convert( int , @vPrintEndScore )
    set @fScore = convert( int , @vScore )
    set @iTargetExamineeCategoryID1 = convert( int , @vTargetExamineeCategoryID1 )
    set @iTargetExamineeAdmission1 = convert( int , @vTargetExamineeAdmission1 )
    set @iTargetExamineeSex1 = convert( int , @vTargetExamineeSex1 )
    if @vTargetExamineeCategoryID2 = '' or @vTargetExamineeCategoryID2 = '-1'
        set @iTargetExamineeCategoryID2 = null
    else
      begin
        set @iTargetExamineeCategoryID2 = convert( int , @vTargetExamineeCategoryID2 )
        set @iTargetExamineeAdmission2 = convert( int , @vTargetExamineeAdmission2 )
        set @iTargetExamineeSex2 = convert( int , @vTargetExamineeSex2 )
      end
    if @vTargetExamineeCategoryID3 = '' or @vTargetExamineeCategoryID3 = '-1'
        set @iTargetExamineeCategoryID3 = null
    else
      begin
        set @iTargetExamineeCategoryID3 = convert( int , @vTargetExamineeCategoryID3 )
        set @iTargetExamineeAdmission3 = convert( int , @vTargetExamineeAdmission3 )
        set @iTargetExamineeSex3 = convert( int , @vTargetExamineeSex3 )
      end
create table #wk1(
	iExamineeProfileId	int
)
create table #wk2(
	iExamineeProfileId	int
)
create table #wk3(
	iExamineeProfileId	int
)
create table #wk4(
	iExamineeProfileId	int
,	fScore			float
)
    exec uspSTEGetTargetExaminee	@vNendo	,	@iTargetExamineeCategoryID1	,	@iTargetExamineeAdmission1	,	@iTargetExamineeSex1	,	'#wk1'
    if @iTargetExamineeCategoryID2 is not null
        exec uspSTEGetTargetExaminee	@vNendo	,	@iTargetExamineeCategoryID2	,	@iTargetExamineeAdmission2	,	@iTargetExamineeSex2	,	'#wk2'
    if @iTargetExamineeCategoryID3 is not null
        exec uspSTEGetTargetExaminee	@vNendo	,	@iTargetExamineeCategoryID3	,	@iTargetExamineeAdmission3	,	@iTargetExamineeSex3	,	'#wk3'
    insert into #wk4(iExamineeProfileId) select * from #wk1
    insert into #wk4(iExamineeProfileId) select * from #wk2 as w2 where not exists ( select 1 from #wk4 as w4 where w4.iExamineeProfileId = w2.iExamineeProfileId )
    insert into #wk4(iExamineeProfileId) select * from #wk3 as w3 where not exists ( select 1 from #wk4 as w4 where w4.iExamineeProfileId = w3.iExamineeProfileId )
create table #wkScore(
	fStartScore	float
,	fEndScore	float
)
    set @fStartScore = @iPrintStartScore
    while @fStartScore < @iPrintEndScore
      begin
        insert into #wkScore select @fStartScore , case when @fStartScore + @fScore >= @iPrintEndScore then 401 else @fStartScore + @fScore end
        set @fStartScore = @fStartScore + @fScore
      end
declare
	@vTotalCondition	varchar(256)
,	@vExamineeCondition	varchar(256)
,	@SQL		nvarchar(1024)
    select @vTotalCondition = vCondition
    from tbSTEtotalCategory
    where iTotalCategoryID = @iTotalCategoryID
--    select @vExamineeCondition = vCondition
--    from tbSTEExamineeCategory
--    where iExamineeCategoryID = 1
    set @SQL = 'update #wk4'
    set @SQL = @SQL + ' set fScore = ( ' 
    set @SQL = @SQL + ' select ' + @vTotalCondition
    set @SQL = @SQL + ' from #wk4 as ep'
--    set @SQL = @SQL + 'inner join tbSTEExamineeProfile as ep on t1.iExamineeProfileID = ep.iExamineeProfileID'
    set @SQL = @SQL + ' inner join tbSTEScoreProfile as sc on sc.iExamineeProfileID = ep.iExamineeProfileID  '
--    set @SQL = @SQL + ' inner join tbSTEScoreDetail as sd on sd.iScoreProfileID = sc.iScoreProfileID'
    set @SQL = @SQL + ' inner join tbSTESubjectProfile as sp on sp.iSubjectProfileID = sc.iSubjectProfileID'
    set @SQL = @SQL + ' where ep.iExamineeProfileID = #wk4.iExamineeProfileID'
--    if @vExamineeCondition <> ''
--        set @SQL = @SQL + ' and ' + @vExamineeCondition
    set @SQL = @SQL + ' group by'
    set @SQL = @SQL + '  ep.iExamineeProfileID'
    set @SQL = @SQL + ' )'
--print @SQL
    exec sp_executesql @SQL

    select @lCnt1 = count(*) from #wk1
    select @lCnt2 = count(*) from #wk2
    select @lCnt3 = count(*) from #wk3

    select	@fMax1 = max( case when t1.iExamineeProfileId is not null then ep.fScore else null end )
	,	@fMax2 = max( case when t2.iExamineeProfileId is not null then ep.fScore else null end )
	,	@fMax3 = max( case when t3.iExamineeProfileId is not null then ep.fScore else null end )
	,	@fMin1 = min( case when t1.iExamineeProfileId is not null then ep.fScore else null end )
	,	@fMin2 = min( case when t2.iExamineeProfileId is not null then ep.fScore else null end )
	,	@fMin3 = min( case when t3.iExamineeProfileId is not null then ep.fScore else null end )
	,	@fSum1 = sum( case when t1.iExamineeProfileId is not null then ep.fScore else null end )
	,	@fSum2 = sum( case when t2.iExamineeProfileId is not null then ep.fScore else null end )
	,	@fSum3 = sum( case when t3.iExamineeProfileId is not null then ep.fScore else null end )
	,	@fAvg1 = avg( case when t1.iExamineeProfileId is not null then ep.fScore else null end )
	,	@fAvg2 = avg( case when t2.iExamineeProfileId is not null then ep.fScore else null end )
	,	@fAvg3 = avg( case when t3.iExamineeProfileId is not null then ep.fScore else null end )
    from 		#wk4 as ep
    left outer join	#wk1 as t1 on t1.iExamineeProfileId = ep.iExamineeProfileId
    left outer join	#wk2 as t2 on t2.iExamineeProfileId = ep.iExamineeProfileId
    left outer join	#wk3 as t3 on t3.iExamineeProfileId = ep.iExamineeProfileId
/*
    select
		@fSd1 = SQRT( sum( case when t1.iExamineeProfileId is not null then POWER ( ep.fScore - @fAvg1 , 2 ) else null end ) / @lCnt1)
	,	@fSd2 = SQRT( sum( case when t2.iExamineeProfileId is not null then POWER ( ep.fScore - @fAvg2 , 2 ) else null end ) / @lCnt2)
	,	@fSd3 = SQRT( sum( case when t3.iExamineeProfileId is not null then POWER ( ep.fScore - @fAvg3 , 2 ) else null end ) / @lCnt3)
    from 		#wk4 as ep
    left outer join	#wk1 as t1 on t1.iExamineeProfileId = ep.iExamineeProfileId
    left outer join	#wk2 as t2 on t2.iExamineeProfileId = ep.iExamineeProfileId
    left outer join	#wk3 as t3 on t3.iExamineeProfileId = ep.iExamineeProfileId
*/
    select
		@fSd1	=	STDEVP(	ep.fScore	)
    from 		#wk4 as ep
    inner join		#wk1 as t1 on t1.iExamineeProfileId = ep.iExamineeProfileId
    select
		@fSd2	=	STDEVP(	ep.fScore	)
    from 		#wk4 as ep
    inner join		#wk2 as t1 on t1.iExamineeProfileId = ep.iExamineeProfileId
    select
		@fSd3	=	STDEVP(	ep.fScore	)
    from 		#wk4 as ep
    inner join		#wk3 as t1 on t1.iExamineeProfileId = ep.iExamineeProfileId

    SELECT
		ws.fStartScore as fStartScore
	,	case when ws.fEndScore > @iPrintEndScore then @iPrintEndScore else ws.fEndScore end as fEndScore
	,	sum( case when t1.iExamineeProfileId is not null and ep.fScore >= ws.fStartScore and ep.fScore < ws.fEndScore then 1 else 0 end ) as lCnt1
	,	sum( case when t2.iExamineeProfileId is not null and ep.fScore >= ws.fStartScore and ep.fScore < ws.fEndScore then 1 else 0 end ) as lCnt2
	,	sum( case when t3.iExamineeProfileId is not null and ep.fScore >= ws.fStartScore and ep.fScore < ws.fEndScore then 1 else 0 end ) as lCnt3
	,	@fMax1	as	fMax1
	,	@fMax2	as	fMax2
	,	@fMax3	as	fMax3
	,	@fMin1	as	fMin1
	,	@fMin2	as	fMin2
	,	@fMin3	as	fMin3
	,	@fAvg1	as	fAvg1
	,	@fAvg2	as	fAvg2
	,	@fAvg3	as	fAvg3
	,	@fSd1	as	fSd1
	,	@fSd2	as	fSd2
	,	@fSd3	as	fSd3
	,	@fSum1	as	fSum1
	,	@fSum2	as	fSum2
	,	@fSum3	as	fSum3
	,	sum( case when t1.iExamineeProfileId is not null and ep.fScore < ws.fEndScore then 1 else 0 end ) as lRuiCnt1
	,	sum( case when t2.iExamineeProfileId is not null and ep.fScore < ws.fEndScore then 1 else 0 end ) as lRuiCnt2
	,	sum( case when t3.iExamineeProfileId is not null and ep.fScore < ws.fEndScore then 1 else 0 end ) as lRuiCnt3
    FROM	#wk4 as ep
    left outer join	#wk1 as t1 on t1.iExamineeProfileId = ep.iExamineeProfileId
    left outer join	#wk2 as t2 on t2.iExamineeProfileId = ep.iExamineeProfileId
    left outer join	#wk3 as t3 on t3.iExamineeProfileId = ep.iExamineeProfileId
    ,		#wkScore as ws
    group by 	ws.fStartScore , ws.fEndScore
    order by	ws.fStartScore DESC
drop table #wk1
drop table #wk2
drop table #wk3
drop table #wk4
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

