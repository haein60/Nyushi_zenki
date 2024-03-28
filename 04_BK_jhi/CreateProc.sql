if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UspCTMAllocateDay]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[UspCTMAllocateDay]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UspCTMAllocateDay1Room]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[UspCTMAllocateDay1Room]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UspCTMAllocateDay2Room]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[UspCTMAllocateDay2Room]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UspCTMAllocateDay3Room]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[UspCTMAllocateDay3Room]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UspCTMAllocateRoom]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[UspCTMAllocateRoom]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UspCTMAllocateRoomDay1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[UspCTMAllocateRoomDay1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UspCTMAllocateRoomDay2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[UspCTMAllocateRoomDay2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UspCTMAllocateRoomDay3]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[UspCTMAllocateRoomDay3]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UspCTMCalScore]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[UspCTMCalScore]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[uspSTEConvertExaminee]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[uspSTEConvertExaminee]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[uspSTEConvertModify]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[uspSTEConvertModify]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[uspSTEConvertSchool]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[uspSTEConvertSchool]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[uspSTEConvertScore]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[uspSTEConvertScore]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[uspSTEInsertExaminee]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[uspSTEInsertExaminee]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[uspSTESeisekiIchiran]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[uspSTESeisekiIchiran]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[uspSTESeisekiStudentScore]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[uspSTESeisekiStudentScore]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[uspSTRSeisekiStudentScore]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[uspSTRSeisekiStudentScore]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[uspSTRWatchReport]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[uspSTRWatchReport]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbCpfSystemUser]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbCpfSystemUser]
GO

CREATE TABLE [dbo].[tbCpfSystemUser] (
	[iUserID] [int] NOT NULL ,
	[vLoginID] [varchar] (16) COLLATE Japanese_CI_AS NULL ,
	[vPassword] [varchar] (256) COLLATE Japanese_CI_AS NULL ,
	[iUserLevel] [int] NULL 
) ON [PRIMARY]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- ****************************************************************************************************************************
-- Stored Procedure		: UspCTMAllocateRoom
-- Input Parametrs		: Nil
-- Created					: 22/10/2001		
-- Author					: Dileep Cherian
-- Output					: Recordsets
-- Modification History	: 
-- Reference				: Functional Spec of Distribution Of Examinee.doc (ver 1.0)
-- ****************************************************************************************************************************
CREATE PROCEDURE UspCTMAllocateDay
AS
BEGIN
SET NOCOUNT ON
DECLARE
@iExamineeId INTEGER,
@iSex INTEGER,
@iDay1Flag INTEGER,
@iDay2Flag INTEGER,
@iDay3Flag INTEGER,
@iMultipleFlag INTEGER,
@iHighSchoolId INTEGER,
@iNoOfExamineeDay1 INTEGER,
@iNoOfExamineeDay2 INTEGER,
@iNoOfExamineeDay3 INTEGER,
@iCountDay1 INTEGER,
@iCountDay2 INTEGER,
@iCountDay3 INTEGER,
@iNoOfConditions INTEGER,
@MultipleValue INTEGER;
/* Store tbSTESecondExamProfile table into this tempporary table */
CREATE TABLE #TempSecondExam
(
	dtDay1 DATETIME,
	dtDay2 DATETIME,
	dtDay3 DATETIME,
	iNoExamineeDay1 INTEGER,
	iNoExamineeDay2 INTEGER,
	iNoExamineeDay3 INTEGER,
	iNoRoomDay1 INTEGER,
	iNoRoomDay2 INTEGER,
	iNoRoomDay3 INTEGER
)
INSERT INTO #TempSecondExam 
SELECT dtSecondExamDay1, dtSecondExamDay2, dtSecondExamDay3,
iNumberOfExamineeDay1, iNumberOfExamineeDay2, iNumberOfExamineeDay3,
iNumberOfRoomDay1, iNumberOfRoomDay2, iNumberOfRoomDay3 FROM tbSTESecondExamProfile 
WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile WHERE iActiveFlag=1)
SELECT @iNoOfExamineeDay1 = (SELECT iNoExamineeDay1 FROM #TempSecondExam)
SELECT @iNoOfExamineeDay2 = (SELECT iNoExamineeDay2 FROM #TempSecondExam)
SELECT @iNoOfExamineeDay3 = (SELECT iNoExamineeDay3 FROM #TempSecondExam)
CREATE TABLE #TempDay1
(
	iExamineeProfileId INTEGER,	
	iSex INTEGER,
	iPreferenceDay1Flag INTEGER,
	iPreferenceDay2Flag INTEGER,
	iPreferenceDay3Flag INTEGER,
	iMultipleApplyFlag INTEGER,
	iHighSchoolId INTEGER
)
CREATE TABLE #TempDay2
(
	iExamineeProfileId INTEGER,	
	iSex INTEGER,
	iPreferenceDay1Flag INTEGER,
	iPreferenceDay2Flag INTEGER,
	iPreferenceDay3Flag INTEGER,
	iMultipleApplyFlag INTEGER,
	iHighSchoolId INTEGER
)
CREATE TABLE #TempDay3
(
	iExamineeProfileId INTEGER,	
	iSex INTEGER,
	iPreferenceDay1Flag INTEGER,
	iPreferenceDay2Flag INTEGER,
	iPreferenceDay3Flag INTEGER,
	iMultipleApplyFlag INTEGER,
	iHighSchoolId INTEGER
)
CREATE TABLE #Day1Excess
(
	iExamineeProfileId INTEGER,	
	iSex INTEGER,
	iPreferenceDay1Flag INTEGER,
	iPreferenceDay2Flag INTEGER,
	iPreferenceDay3Flag INTEGER,
	iMultipleApplyFlag INTEGER,
	iHighSchoolId INTEGER
)
CREATE TABLE #Day2Excess
(
	iExamineeProfileId INTEGER,	
	iSex INTEGER,
	iPreferenceDay1Flag INTEGER,
	iPreferenceDay2Flag INTEGER,
	iPreferenceDay3Flag INTEGER,
	iMultipleApplyFlag INTEGER,
	iHighSchoolId INTEGER
)
CREATE TABLE #Day3Excess
(
	iExamineeProfileId INTEGER,	
	iSex INTEGER,
	iPreferenceDay1Flag INTEGER,
	iPreferenceDay2Flag INTEGER,
	iPreferenceDay3Flag INTEGER,
	iMultipleApplyFlag INTEGER,
	iHighSchoolId INTEGER
)
CREATE TABLE #Day12Excess
(
	iExamineeProfileId INTEGER,	
	iSex INTEGER,
	iPreferenceDay1Flag INTEGER,
	iPreferenceDay2Flag INTEGER,
	iPreferenceDay3Flag INTEGER,
	iMultipleApplyFlag INTEGER,
	iHighSchoolId INTEGER
)
CREATE TABLE #Day13Excess
(
	iExamineeProfileId INTEGER,	
	iSex INTEGER,
	iPreferenceDay1Flag INTEGER,
	iPreferenceDay2Flag INTEGER,
	iPreferenceDay3Flag INTEGER,
	iMultipleApplyFlag INTEGER,
	iHighSchoolId INTEGER
)
CREATE TABLE #Day23Excess
(
	iExamineeProfileId INTEGER,	
	iSex INTEGER,
	iPreferenceDay1Flag INTEGER,
	iPreferenceDay2Flag INTEGER,
	iPreferenceDay3Flag INTEGER,
	iMultipleApplyFlag INTEGER,
	iHighSchoolId INTEGER
)
CREATE TABLE #TempDay12
(
	iExamineeProfileId INTEGER,	
	iSex INTEGER,
	iPreferenceDay1Flag INTEGER,
	iPreferenceDay2Flag INTEGER,
	iPreferenceDay3Flag INTEGER,
	iMultipleApplyFlag INTEGER,
	iHighSchoolId INTEGER
)
CREATE TABLE #TempDay13
(
	iExamineeProfileId INTEGER,	
	iSex INTEGER,
	iPreferenceDay1Flag INTEGER,
	iPreferenceDay2Flag INTEGER,
	iPreferenceDay3Flag INTEGER,
	iMultipleApplyFlag INTEGER,
	iHighSchoolId INTEGER
)
CREATE TABLE #TempDay23
(
	iExamineeProfileId INTEGER,	
	iSex INTEGER,
	iPreferenceDay1Flag INTEGER,
	iPreferenceDay2Flag INTEGER,
	iPreferenceDay3Flag INTEGER,
	iMultipleApplyFlag INTEGER,
	iHighSchoolId INTEGER
)
CREATE TABLE #TempDay123
(
	iExamineeProfileId INTEGER,	
	iSex INTEGER,
	iPreferenceDay1Flag INTEGER,
	iPreferenceDay2Flag INTEGER,
	iPreferenceDay3Flag INTEGER,
	iMultipleApplyFlag INTEGER,
	iHighSchoolId INTEGER
)
SET @iCountDay1 = 1
SET @iCountDay2 = 1
SET @iCountDay3 = 1
SET @iNoOfConditions = 1
WHILE @iNoOfConditions <= 3
BEGIN
	DECLARE ExamineeCursor CURSOR FOR
	SELECT iExamineeProfileId, iSex, iPreferenceDay1Flag, iPreferenceDay2Flag, iPreferenceDay3Flag, iMultipleApplyFlag, iHighSchoolId
	FROM tbSTEExamineeProfile
	WHERE iNendo=(SELECT iNendo FROM tbSTESystemProfile WHERE iActiveFlag=1)
	AND iExamineeStatus = 1 AND iAbsentFlag = 0
	ORDER BY iExamineeProfileId
	
	OPEN ExamineeCursor
	
	FETCH NEXT FROM ExamineeCursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
	WHILE @@FETCH_STATUS = 0
	BEGIN  
		IF @iNoOfConditions = 1 OR @iNoOfConditions = 2
	   BEGIN
			IF @iNoOfConditions = 1
			   SET @MultipleValue = 1
			ELSE
		   	SET @MultipleValue = 0  	 
	    
			IF @iDay1Flag = 1 AND @iDay2Flag = 0 AND @iDay3Flag = 0 AND @iMultipleFlag = @MultipleValue
			BEGIN				
				IF @iCountDay1 <= @iNoOfExamineeDay1
				BEGIN
					SET @iCountDay1 = @iCountDay1 + 1
					INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
				END
				ELSE
					INSERT INTO #Day1Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
			END	 
	    
			IF @iDay1Flag = 0 AND @iDay2Flag = 1 AND @iDay3Flag = 0 AND @iMultipleFlag = @MultipleValue
			BEGIN				
				IF @iCountDay2 <= @iNoOfExamineeDay2
				BEGIN
					SET @iCountDay2 = @iCountDay2 + 1
					INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
				END
				ELSE
					INSERT INTO #Day2Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
			END 
			IF @iDay1Flag = 0 AND @iDay2Flag = 0 AND @iDay3Flag = 1 AND @iMultipleFlag = @MultipleValue
			BEGIN				
				IF @iCountDay3 <= @iNoOfExamineeDay3
				BEGIN
					SET @iCountDay3 = @iCountDay3 + 1
					INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)	
				END
				ELSE
					INSERT INTO #Day3Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
			END 
		END     
	
	   ELSE	/* condition <> 1 or 2 */
			BEGIN
				IF @iNoOfConditions = 3
				BEGIN
					IF @iDay1Flag = 1 AND @iDay2Flag = 1 AND @iDay3Flag = 0
						INSERT INTO #TempDay12 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)		
					
					IF @iDay1Flag = 1 AND @iDay2Flag = 0 AND @iDay3Flag = 1
						INSERT INTO #TempDay13 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)		
					
					IF @iDay1Flag = 0 AND @iDay2Flag = 1 AND @iDay3Flag = 1
						INSERT INTO #TempDay23 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)		
					
					IF @iDay1Flag = 1 AND @iDay2Flag = 1 AND @iDay3Flag = 1
						INSERT INTO #TempDay123 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)		
				END
			END
		FETCH NEXT FROM ExamineeCursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
	END	/* End While - Cursor */
	
	CLOSE ExamineeCursor
	DEALLOCATE ExamineeCursor
	SET @iNoOfConditions = @iNoOfConditions + 1
END	/* End While - @iNoOfConditions */
/* Insert the records of #TempDay12 */
DECLARE temp_cursor CURSOR FOR		
SELECT * FROM #TempDay12
OPEN temp_cursor
FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
WHILE @@FETCH_STATUS = 0
BEGIN    
	
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	BEGIN
		SET @iCountDay1 = @iCountDay1 + 1
		INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	END
	ELSE	/* Insert into Day 3 */
	BEGIN		
		IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
		BEGIN
			SET @iCountDay2 = @iCountDay2 + 1
			INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
		END
		ELSE	/* Insert into Day 3 */
		  INSERT INTO #Day12Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)		
	END
	FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
END
CLOSE temp_cursor
DEALLOCATE temp_cursor
/* Insert the records of #TempDay13 */
DECLARE temp_cursor CURSOR FOR		
SELECT * FROM #TempDay13
OPEN temp_cursor
FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
WHILE @@FETCH_STATUS = 0
BEGIN    	
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	BEGIN
		 SET @iCountDay1 = @iCountDay1 + 1
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	END
	ELSE	/* Insert into Day 3 */
	BEGIN		
		IF @iCountDay3 <= @iNoOfExamineeDay3	/* Insert into Day 3 */
		BEGIN
			SET @iCountDay3 = @iCountDay3 + 1
			INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
		END
		ELSE	/* Insert into Day 2 */
		  INSERT INTO #Day13Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	END
	FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
END  
CLOSE temp_cursor
DEALLOCATE temp_cursor   
/* Insert the records of #TempDay23 */
DECLARE temp_cursor CURSOR FOR		
SELECT * FROM #TempDay23
OPEN temp_cursor
FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
WHILE @@FETCH_STATUS = 0
BEGIN    	
	IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
	BEGIN		
		 SET @iCountDay2 = @iCountDay2 + 1
	    INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)		
	END
	ELSE	/* Insert into Day 3 */
	BEGIN		 
		 IF @iCountDay3 <= @iNoOfExamineeDay3	/* Insert into Day 3 */
		 BEGIN
			SET @iCountDay3 = @iCountDay3 + 1
		 	INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
		 END
		 ELSE	/* Insert into Day 1 */
			INSERT INTO #Day23Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	END
	FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
END     
CLOSE temp_cursor
DEALLOCATE temp_cursor
/* Insert the records of #TempDay123 */
DECLARE temp_cursor CURSOR FOR		
SELECT * FROM #TempDay123
OPEN temp_cursor
FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
WHILE @@FETCH_STATUS = 0
BEGIN    	
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	BEGIN
		 SET @iCountDay1 = @iCountDay1 + 1
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	END
	ELSE	/* Insert into Day 2 */
	BEGIN		
		IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
		BEGIN
			SET @iCountDay2 = @iCountDay2 + 1
			INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
		END
		ELSE	/* Insert into Day 3 */
		BEGIN
			SET @iCountDay3 = @iCountDay3 + 1
			INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)                                
		END
	END
	FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
END
CLOSE temp_cursor
DEALLOCATE temp_cursor
/* Insert the records of #Day1Excess */
DECLARE temp_cursor CURSOR FOR		
SELECT * FROM #Day1Excess
OPEN temp_cursor
FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
WHILE @@FETCH_STATUS = 0
BEGIN    	
	IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
	BEGIN
		 SET @iCountDay2 = @iCountDay2 + 1
	    INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	END
	ELSE	/* Insert into Day 3 */
	BEGIN	    
	    SET @iCountDay3 = @iCountDay3 + 1
       INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)                                
	END
	FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
END
CLOSE temp_cursor
DEALLOCATE temp_cursor
/* Insert the records of #Day2Excess */
DECLARE temp_cursor CURSOR FOR		
SELECT * FROM #Day2Excess
OPEN temp_cursor
FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
WHILE @@FETCH_STATUS = 0
BEGIN    
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	BEGIN
		 	SET @iCountDay1 = @iCountDay1 + 1
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	END
	ELSE	/* Insert into Day 3 */
	BEGIN	    	   
	    SET @iCountDay3 = @iCountDay3 + 1
       INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)                                
	END
	FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
END
CLOSE temp_cursor
DEALLOCATE temp_cursor
/* Insert the records of #Day3Excess */
DECLARE temp_cursor CURSOR FOR		
SELECT * FROM #Day3Excess
OPEN temp_cursor
FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
WHILE @@FETCH_STATUS = 0
BEGIN    
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	BEGIN
       SET @iCountDay1 = @iCountDay1 + 1
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	END
	ELSE	/* Insert into Day 2 */
	BEGIN	    	   
	    SET @iCountDay2 = @iCountDay2 + 1
       INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)                                
	END
	FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
END
CLOSE temp_cursor
DEALLOCATE temp_cursor
/* Insert the records of #Day12Excess */
DECLARE temp_cursor CURSOR FOR		
SELECT * FROM #Day12Excess
OPEN temp_cursor
FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
WHILE @@FETCH_STATUS = 0
BEGIN    
	SET @iCountDay3 = @iCountDay3 + 1
	/* Insert into Day 3 */
	INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	
	FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
END
CLOSE temp_cursor
DEALLOCATE temp_cursor
/* Insert the records of #Day13Excess */
DECLARE temp_cursor CURSOR FOR		
SELECT * FROM #Day13Excess
OPEN temp_cursor
FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
WHILE @@FETCH_STATUS = 0
BEGIN    
	SET @iCountDay2 = @iCountDay2 + 1
	/* Insert into Day 2 */
   INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	
	FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
END
CLOSE temp_cursor
DEALLOCATE temp_cursor
/* Insert the records of #Day23Excess */
DECLARE temp_cursor CURSOR FOR		
SELECT * FROM #Day23Excess
OPEN temp_cursor
FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
WHILE @@FETCH_STATUS = 0
BEGIN    
	SET @iCountDay1 = @iCountDay1 + 1
	/* Insert into Day 1 */
   INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	
	FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
END
CLOSE temp_cursor
DEALLOCATE temp_cursor

exec UspCTMAllocateDay1Room
exec UspCTMAllocateDay2Room
exec UspCTMAllocateDay3Room

delete from tbSTEExamineeRoomProfile where exists ( select 1 from tbSTEexamineeProfile as ep where ep.iNendo = ( select top 1 iNendo from tbSTEsystemProfile where iActiveFlag = 1 )  and ep.iExamineeprofileid = tbSTEExamineeRoomProfile.iExamineeprofileid )

--SELECT * FROM #Day1Excess
--SELECT * FROM #Day2Excess
--SELECT * FROM #Day3Excess
--SELECT * FROM #Day12Excess
--SELECT * FROM #Day13Excess
--SELECT * FROM #Day23Excess
END	/* Final End */
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

-- ****************************************************************************************************************************
-- Stored Procedure		: UspCTMAllocateDay1Room
-- Input Parametrs		: Nil
-- Created					: 22/10/2001		
-- Author					: Dileep Cherian
--	Output					: Nil
--	Modification History	: Nil
-- Reference				: Functional Spec of Distribution Of Examinee.doc (ver 1.0)
-- ****************************************************************************************************************************
CREATE PROCEDURE UspCTMAllocateDay1Room
AS
BEGIN
DECLARE
@iExamineeId INTEGER,			
@iSex INTEGER,
@iHighSchoolId INTEGER,
@iTotalExamineeDay1 INTEGER,		-- total examinees allocated for day1
@iTotalRoomsDay1 INTEGER,			-- total rooms allocated for day1
@Capacity INTEGER,					-- room capacity
@Id INTEGER,							-- examinee id
@counter INTEGER,
@SchoolId INTEGER,
@iRoomId INTEGER,
@dtExamDay1 DATETIME,				-- actual date of day1 exam
@iCapacity INTEGER,
@iCount INTEGER,
@iCounter INTEGER,
@iNewId INTEGER,
@iNormalInterview INTEGER,			
@iTotalRoomsAvailable INTEGER;	-- total available rooms
-- store the details of all available rooms in temporary table
CREATE TABLE #RoomDetail			
(
	iRoomId INTEGER,
	iRoomProfileId INTEGER,
	iMaxCapacity INTEGER
)
-- temporary allocation table
CREATE TABLE #TempRoom1
(
	iRoomId INTEGER,
	iExamineeId INTEGER,
	iSex INTEGER,	
	iHighSchoolId INTEGER
)
-- get these values from UspCTMAllocateRoom
SET @iTotalExamineeDay1 = (SELECT COUNT(*) FROM #TempDay1)
SET @iTotalRoomsDay1 = (SELECT iNoRoomDay1 FROM #TempSecondExam)
SET @counter = 1
-- populate the #RoomDetail table
DECLARE temp_cursor2 CURSOR FOR
-- select only those rooms which are flagged for Interviews
SELECT iRoomProfileId, iMaxCapacity FROM tbSTERoomProfile 
WHERE iMaxCapacity > 0 AND iInterviewRoomFlag = 0 ORDER BY iRoomProfileId
OPEN temp_cursor2
FETCH NEXT FROM temp_cursor2 INTO @Id, @Capacity
WHILE @@FETCH_STATUS = 0 AND @counter <= @iTotalRoomsDay1
BEGIN
	INSERT INTO #RoomDetail VALUES(@counter, @Id, @Capacity)
	SET @counter = @counter + 1
	FETCH NEXT FROM temp_cursor2 INTO @Id, @Capacity
END
CLOSE temp_cursor2
DEALLOCATE temp_cursor2
-- get the number of available rooms
SELECT @iTotalRoomsAvailable = (SELECT COUNT(*) FROM #RoomDetail)
SET @iRoomId = 1
-- loop through all the distinct schools ids - outer loop
DECLARE temp_cursor3 CURSOR FOR
SELECT DISTINCT iHighSchoolId FROM #TempDay1
OPEN temp_cursor3
FETCH NEXT FROM temp_cursor3 INTO @SchoolId
WHILE @@FETCH_STATUS = 0 
BEGIN		
	-- loop through all the male examinees in that school - inner loop (male)
	DECLARE temp_cursor4 CURSOR FOR
	SELECT iExamineeProfileId, iSex, iHighSchoolId FROM #TempDay1 WHERE iSex = 0 AND iHighSchoolId = @SchoolId
	OPEN temp_cursor4
	FETCH NEXT FROM temp_cursor4 INTO @iExamineeId, @iSex, @iHighSchoolId
	WHILE @@FETCH_STATUS = 0 
	BEGIN		
		SET @iCounter = 1
		WHILE @iCounter <= @iTotalRoomsDay1  -- do this till the number of rooms allocated for the day is not exceeded
		BEGIN
			SELECT @iCount = (SELECT COUNT(*) FROM #TempRoom1 WHERE iRoomId = @iRoomId)
			SELECT @iCapacity = (SELECT iMaxCapacity FROM #RoomDetail WHERE iRoomId = @iRoomId)
			
			IF @iCount < @iCapacity	-- check for the capacity of the room
			BEGIN	
				INSERT INTO #TempRoom1	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay1
					SET @iRoomId = 1
				BREAK;
			END
			ELSE	-- capacity of room room is full, so move to the enxt room
			BEGIN
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay1
					-- move to the first room, after reaching the last room
					SET @iRoomId = 1
				INSERT INTO #TempRoom1	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay1
					SET @iRoomId = 1
				BREAK;
			END
			SET @iCounter = @iCounter + 1
		END
		FETCH NEXT FROM temp_cursor4 INTO @iExamineeId, @iSex, @iHighSchoolId
	END
	CLOSE temp_cursor4
	DEALLOCATE temp_cursor4 
	-- loop through all the female examinees in that school - inner loop(female)
	DECLARE temp_cursor5 CURSOR FOR
	SELECT iExamineeProfileId, iSex, iHighSchoolId FROM #TempDay1 WHERE iSex = 1 AND iHighSchoolId = @SchoolId
	OPEN temp_cursor5
	FETCH NEXT FROM temp_cursor5 INTO @iExamineeId, @iSex, @iHighSchoolId
	WHILE @@FETCH_STATUS = 0 
	BEGIN
		SET @iCounter = 1
		WHILE @iCounter <= @iTotalRoomsDay1
		BEGIN
			SELECT @iCount = (SELECT COUNT(*) FROM #TempRoom1 WHERE iRoomId = @iRoomId)
			SELECT @iCapacity = (SELECT iMaxCapacity FROM #RoomDetail WHERE iRoomId = @iRoomId)
			IF @iCount < @iCapacity		-- check for the capacity of the room
			BEGIN	
				INSERT INTO #TempRoom1	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay1
					SET @iRoomId = 1
				BREAK;
			END
			ELSE		-- capacity of room room is full, so move to the next room
			BEGIN
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay1
					-- move to the first room, after reaching the last room
					SET @iRoomId = 1
				INSERT INTO #TempRoom1	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay1
					SET @iRoomId = 1
				BREAK;
			END
			SET @iCounter = @iCounter + 1
		END
		FETCH NEXT FROM temp_cursor5 INTO @iExamineeId, @iSex, @iHighSchoolId
	END
	CLOSE temp_cursor5
	DEALLOCATE temp_cursor5 
	FETCH NEXT FROM temp_cursor3 INTO @SchoolId
END
CLOSE temp_cursor3
DEALLOCATE temp_cursor3
-- actual insertion into tbSTEExamineeProfile and tbSTEExamineeRoomProfile table from #TempRoom1
DECLARE temp_cursor6 CURSOR FOR
SELECT iRoomId, iExamineeId FROM #TempRoom1
OPEN temp_cursor6
FETCH NEXT FROM temp_cursor6 INTO @iRoomId, @iExamineeId
WHILE @@FETCH_STATUS = 0 
BEGIN
	
	SELECT @Id = (SELECT iRoomProfileId FROM #RoomDetail WHERE iRoomId = @iRoomId)	-- get the actual roomprofileid
	SELECT @dtExamDay1 = (SELECT dtDay1 FROM #TempSecondExam)								-- get the exam date for day1
	
	-- update dtSecondExamDay field of tbSTEExamineeProfile for the particular examinee
	UPDATE tbSTEExamineeProfile SET dtSecondExamDay = @dtExamDay1 
	WHERE iExamineeProfileId = @iExamineeId
	
	-- get the subjectid for the normal interview
	SELECT @iNormalInterview = (SELECT iSubjectProfileId FROM tbSTESubjectProfile WHERE iExamType = 2)
	-- delete any existing row from tbSTEExamineeRoomProfile for the selected examinee-subject combination
	DELETE FROM tbSTEExamineeRoomProfile 
	WHERE iExamineeProfileId = @iExamineeId
	AND iSubjectProfileId = @iNormalInterview 
	
	-- get the new id for the tbSTEExamineeRoomProfile table
	SELECT @iNewId = (SELECT MAX(iExamineeRoomProfileId) FROM tbSTEExamineeRoomProfile)
	IF @iNewId IS NULL
		SELECT @iNewId = (SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName = 'tbSTEExamineeRoomProfile')		
	ELSE
		SET @iNewId = @iNewId + 1
	
	-- insert the examinee data for nomal interview into tbSTEEXamineeRoomProfile for the selected examinee
	INSERT INTO tbSTEExamineeRoomProfile VALUES(@iNewId, @iExamineeId, @Id, @iNormalInterview, getdate(),getdate())
		
	FETCH NEXT FROM temp_cursor6 INTO @iRoomId, @iExamineeId
END
CLOSE temp_cursor6
DEALLOCATE temp_cursor6
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

-- ****************************************************************************************************************************
-- Stored Procedure		: UspCTMAllocateDay2Room
-- Input Parametrs		: Nil
-- Created			: 22/10/2001		
-- Author			: Dileep Cherian
--Output				: Nil
--Modification History		: Nil
-- Reference			: Functional Spec of Distribution Of Examinee.doc (ver 1.0)
-- ****************************************************************************************************************************
CREATE PROCEDURE UspCTMAllocateDay2Room
AS
BEGIN
DECLARE
@iExamineeId INTEGER,
@iSex INTEGER,
@iHighSchoolId INTEGER,
@iTotalExamineeDay2 INTEGER,		-- total examinees allocated for day2
@iTotalRoomsDay2 INTEGER,			-- total rooms allocated for day1
@Capacity INTEGER,					-- room capacity
@Id INTEGER,							-- examinee id
@counter INTEGER,
@SchoolId INTEGER,
@iRoomId INTEGER,
@dtExamDay2 DATETIME,				-- actual date of day1 exam
@iCapacity INTEGER,
@iCount INTEGER,
@iCounter INTEGER,
@iNewId INTEGER,
@iNormalInterview INTEGER,
@iTotalRoomsAvailable INTEGER;	-- total available rooms
-- store the details of all available rooms in temporary table
CREATE TABLE #RoomDetail
(
	iRoomId INTEGER,
	iRoomProfileId INTEGER,
	iMaxCapacity INTEGER
)
-- temporary allocation table
CREATE TABLE #TempRoom2
(
	iRoomId INTEGER,
	iExamineeId INTEGER,
	iSex INTEGER,	
	iHighSchoolId INTEGER
)
-- get these values from UspCTMAllocateRoom
SET @iTotalExamineeDay2 = (SELECT COUNT(*) FROM #TempDay2)
SET @iTotalRoomsDay2 = (SELECT iNoRoomDay2 FROM #TempSecondExam)
SET @counter = 1
-- populate the #RoomDetail table
DECLARE temp_cursor2 CURSOR FOR
-- select only those rooms which are flagged for Interviews
SELECT iRoomProfileId, iMaxCapacity FROM tbSTERoomProfile 
WHERE iMaxCapacity > 0 AND iInterviewRoomFlag = 0 ORDER BY iRoomProfileId
OPEN temp_cursor2
FETCH NEXT FROM temp_cursor2 INTO @Id, @Capacity
WHILE @@FETCH_STATUS = 0 AND @counter <= @iTotalRoomsDay2
BEGIN
	INSERT INTO #RoomDetail VALUES(@counter, @Id, @Capacity)
	SET @counter = @counter + 1
	FETCH NEXT FROM temp_cursor2 INTO @Id, @Capacity
END
CLOSE temp_cursor2
DEALLOCATE temp_cursor2
-- get the number of available rooms
SELECT @iTotalRoomsAvailable = (SELECT COUNT(*) FROM #RoomDetail)
SET @iRoomId = 1
-- loop through all the distinct schools ids - outer loop
DECLARE temp_cursor3 CURSOR FOR
SELECT DISTINCT iHighSchoolId FROM #TempDay2
OPEN temp_cursor3
FETCH NEXT FROM temp_cursor3 INTO @SchoolId
WHILE @@FETCH_STATUS = 0 
BEGIN
	-- loop through all the male examinees in that school - inner loop (male)
	DECLARE temp_cursor4 CURSOR FOR
	SELECT iExamineeProfileId, iSex, iHighSchoolId FROM #TempDay2 WHERE iSex = 0 AND iHighSchoolId = @SchoolId
	OPEN temp_cursor4
	FETCH NEXT FROM temp_cursor4 INTO @iExamineeId, @iSex, @iHighSchoolId
	WHILE @@FETCH_STATUS = 0 
	BEGIN
		SET @iCounter = 1
		WHILE @iCounter <= @iTotalRoomsDay2		-- do this till the number of rooms allocated for the day is not exceeded
		BEGIN
			SELECT @iCount = (SELECT COUNT(*) FROM #TempRoom2 WHERE iRoomId = @iRoomId)
			SELECT @iCapacity = (SELECT iMaxCapacity FROM #RoomDetail WHERE iRoomId = @iRoomId)
			
IF @iCount < @iCapacity		-- check for the capacity of the room
			BEGIN	
				INSERT INTO #TempRoom2	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay2
					SET @iRoomId = 1
				BREAK;
			END
			ELSE		-- capacity of room room is full, so move to the enxt room
			BEGIN
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay2
					-- move to the first room, after reaching the last room
					SET @iRoomId = 1
				INSERT INTO #TempRoom2	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay2
					SET @iRoomId = 1
				BREAK;
			END
			SET @iCounter = @iCounter + 1
		END
		FETCH NEXT FROM temp_cursor4 INTO @iExamineeId, @iSex, @iHighSchoolId
	END
	CLOSE temp_cursor4
	DEALLOCATE temp_cursor4 
	
	-- loop through all the female examinees in that school - inner loop(female)
	DECLARE temp_cursor5 CURSOR FOR
	SELECT iExamineeProfileId, iSex, iHighSchoolId FROM #TempDay2 WHERE iSex = 1 AND iHighSchoolId = @SchoolId
	OPEN temp_cursor5
	FETCH NEXT FROM temp_cursor5 INTO @iExamineeId, @iSex, @iHighSchoolId
	WHILE @@FETCH_STATUS = 0 
	BEGIN
		SET @iCounter = 1
		WHILE @iCounter <= @iTotalRoomsDay2
		BEGIN
			SELECT @iCount = (SELECT COUNT(*) FROM #TempRoom2 WHERE iRoomId = @iRoomId)
			SELECT @iCapacity = (SELECT iMaxCapacity FROM #RoomDetail WHERE iRoomId = @iRoomId)
			IF @iCount < @iCapacity		-- check for the capacity of the room
			BEGIN	
				INSERT INTO #TempRoom2	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay2
					SET @iRoomId = 1
				BREAK;
			END
			ELSE		-- capacity of room room is full, so move to the next room
			BEGIN
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay2
					-- move to the first room, after reaching the last room
					SET @iRoomId = 1
				INSERT INTO #TempRoom2	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay2
					SET @iRoomId = 1
				BREAK;
			END
			SET @iCounter = @iCounter + 1
		END
		FETCH NEXT FROM temp_cursor5 INTO @iExamineeId, @iSex, @iHighSchoolId
	END
	CLOSE temp_cursor5
	DEALLOCATE temp_cursor5 
	FETCH NEXT FROM temp_cursor3 INTO @SchoolId
END
CLOSE temp_cursor3
DEALLOCATE temp_cursor3
-- actual insertion into tbSTEExamineeProfile and tbSTEExamineeRoomProfile table from #TempRoom2
DECLARE temp_cursor6 CURSOR FOR
SELECT iRoomId, iExamineeId, iSex, iHighSchoolId FROM #TempRoom2
OPEN temp_cursor6
FETCH NEXT FROM temp_cursor6 INTO @iRoomId, @iExamineeId, @iSex, @iHighSchoolId
WHILE @@FETCH_STATUS = 0 
BEGIN
	SELECT @Id = (SELECT iRoomProfileId FROM #RoomDetail WHERE iRoomId = @iRoomId)
	SELECT @dtExamDay2 = (SELECT dtDay2 FROM #TempSecondExam)
	
	-- update dtSecondExamDay field of tbSTEExamineeProfile for the particular examinee
	UPDATE tbSTEExamineeProfile SET dtSecondExamDay = @dtExamDay2 
	WHERE iExamineeProfileId = @iExamineeId
	
	-- get the subjectid for the normal interview
	SELECT @iNormalInterview = (SELECT iSubjectProfileId FROM tbSTESubjectProfile WHERE iExamType = 2)
	-- delete any existing row from tbSTEExamineeRoomProfile for the selected examinee-subject combination
	DELETE FROM tbSTEExamineeRoomProfile 
	WHERE iExamineeProfileId = @iExamineeId
	AND iSubjectProfileId = @iNormalInterview 
	-- get the new id for the tbSTEExamineeRoomProfile table
	SELECT @iNewId = (SELECT MAX(iExamineeRoomProfileId) FROM tbSTEExamineeRoomProfile)
	IF @iNewId IS NULL
		SELECT @iNewId = (SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName = 'tbSTEExamineeRoomProfile')
	ELSE
		SET @iNewId = @iNewId + 1
	
	-- insert the examinee data for nomal interview into tbSTEEXamineeRoomProfile for the selected examinee
	INSERT INTO tbSTEExamineeRoomProfile VALUES(@iNewId, @iExamineeId, @Id, @iNormalInterview, getdate(),getdate())
		
	FETCH NEXT FROM temp_cursor6 INTO @iRoomId, @iExamineeId, @iSex, @iHighSchoolId
END
CLOSE temp_cursor6
DEALLOCATE temp_cursor6
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

-- ****************************************************************************************************************************
-- Stored Procedure		: UspCTMAllocateDay3Room
-- Input Parametrs		: Nil
-- Created					: 22/10/2001		
-- Author					: Dileep Cherian
--Output						: Nil
--Modification History	: Nil
-- Reference				: Functional Spec of Distribution Of Examinee.doc (ver 1.0)
-- ****************************************************************************************************************************
CREATE PROCEDURE UspCTMAllocateDay3Room
AS
BEGIN
DECLARE
@iExamineeId INTEGER,
@iSex INTEGER,
@iHighSchoolId INTEGER,
@iTotalExamineeDay3 INTEGER,		-- total examinees allocated for day1
@iTotalRoomsDay3 INTEGER,			-- total rooms allocated for day1
@Capacity INTEGER,					-- room capacity
@Id INTEGER,							-- examinee id
@counter INTEGER,
@SchoolId INTEGER,
@iRoomId INTEGER,
@dtExamDay3 DATETIME,				-- actual date of day3 exam
@iCapacity INTEGER,
@iCount INTEGER,
@iCounter INTEGER,
@iNewId INTEGER,
@iNormalInterview INTEGER,
@iTotalRoomsAvailable INTEGER;	-- total available rooms
-- store the details of all available rooms in temporary table
CREATE TABLE #RoomDetail
(
	iRoomId INTEGER,
	iRoomProfileId INTEGER,
	iMaxCapacity INTEGER
)
-- temporary allocation table
CREATE TABLE #TempRoom3
(
	iRoomId INTEGER,
	iExamineeId INTEGER,
	iSex INTEGER,	
	iHighSchoolId INTEGER
)
-- get these values from UspCTMAllocateRoom
SET @iTotalExamineeDay3 = (SELECT COUNT(*) FROM #TempDay3)
SET @iTotalRoomsDay3 = (SELECT iNoRoomDay3 FROM #TempSecondExam)
SET @counter = 1
-- populate the #RoomDetail table
DECLARE temp_cursor2 CURSOR FOR
-- select only those rooms which are flagged for Interviews
SELECT iRoomProfileId, iMaxCapacity FROM tbSTERoomProfile 
WHERE iMaxCapacity > 0 AND iInterviewRoomFlag = 0 ORDER BY iRoomProfileId
OPEN temp_cursor2
FETCH NEXT FROM temp_cursor2 INTO @Id, @Capacity
WHILE @@FETCH_STATUS = 0 AND @counter <= @iTotalRoomsDay3
BEGIN
	INSERT INTO #RoomDetail VALUES(@counter, @Id, @Capacity)
	SET @counter = @counter + 1
	FETCH NEXT FROM temp_cursor2 INTO @Id, @Capacity
END
CLOSE temp_cursor2
DEALLOCATE temp_cursor2
-- get the number of available rooms
SELECT @iTotalRoomsAvailable = (SELECT COUNT(*) FROM #RoomDetail)
SET @iRoomId = 1
-- loop through all the distinct schools ids - outer loop
DECLARE temp_cursor3 CURSOR FOR
SELECT DISTINCT iHighSchoolId FROM #TempDay3
OPEN temp_cursor3
FETCH NEXT FROM temp_cursor3 INTO @SchoolId
WHILE @@FETCH_STATUS = 0 
BEGIN
	-- loop through all the male examinees in that school - inner loop (male)
	DECLARE temp_cursor4 CURSOR FOR
	SELECT iExamineeProfileId, iSex, iHighSchoolId FROM #TempDay3 WHERE iSex = 0 AND iHighSchoolId = @SchoolId
	OPEN temp_cursor4
	FETCH NEXT FROM temp_cursor4 INTO @iExamineeId, @iSex, @iHighSchoolId
	WHILE @@FETCH_STATUS = 0 
	BEGIN
		SET @iCounter = 1
		WHILE @iCounter <= @iTotalRoomsDay3		-- do this till the number of rooms allocated for the day is not exceeded
		BEGIN
			SELECT @iCount = (SELECT COUNT(*) FROM #TempRoom3 WHERE iRoomId = @iRoomId)
			SELECT @iCapacity = (SELECT iMaxCapacity FROM #RoomDetail WHERE iRoomId = @iRoomId)
			
			IF @iCount < @iCapacity		-- check for the capacity of the room
			BEGIN	
				INSERT INTO #TempRoom3	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay3
					SET @iRoomId = 1
				BREAK;
			END
			ELSE		-- capacity of room room is full, so move to the enxt room
			BEGIN
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay3
					-- move to the first room, after reaching the last room
					SET @iRoomId = 1
				INSERT INTO #TempRoom3	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay3
					SET @iRoomId = 1
				BREAK;
			END
			SET @iCounter = @iCounter + 1
		END
		FETCH NEXT FROM temp_cursor4 INTO @iExamineeId, @iSex, @iHighSchoolId
	END
	CLOSE temp_cursor4
	DEALLOCATE temp_cursor4 
	-- loop through all the female examinees in that school - inner loop(female)
	DECLARE temp_cursor5 CURSOR FOR
	SELECT iExamineeProfileId, iSex, iHighSchoolId FROM #TempDay3 WHERE iSex = 1 AND iHighSchoolId = @SchoolId
	OPEN temp_cursor5
	FETCH NEXT FROM temp_cursor5 INTO @iExamineeId, @iSex, @iHighSchoolId
	WHILE @@FETCH_STATUS = 0 
	BEGIN
		SET @iCounter = 1
		WHILE @iCounter <= @iTotalRoomsDay3
		BEGIN
			SELECT @iCount = (SELECT COUNT(*) FROM #TempRoom3 WHERE iRoomId = @iRoomId)
			SELECT @iCapacity = (SELECT iMaxCapacity FROM #RoomDetail WHERE iRoomId = @iRoomId)
			IF @iCount < @iCapacity		-- check for the capacity of the room
			BEGIN	
				INSERT INTO #TempRoom3	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay3
					SET @iRoomId = 1
				BREAK;
			END
			ELSE		-- capacity of room room is full, so move to the next room
			BEGIN
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay3
					-- move to the first room, after reaching the last room
					SET @iRoomId = 1
				INSERT INTO #TempRoom3	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay3
					SET @iRoomId = 1
				BREAK;
			END
			SET @iCounter = @iCounter + 1
		END
		--PRINT @iRoomId
		FETCH NEXT FROM temp_cursor5 INTO @iExamineeId, @iSex, @iHighSchoolId
	END
	CLOSE temp_cursor5
	DEALLOCATE temp_cursor5 
	FETCH NEXT FROM temp_cursor3 INTO @SchoolId
END
CLOSE temp_cursor3
DEALLOCATE temp_cursor3
-- actual insertion into tbSTEExamineeProfile and tbSTEExamineeRoomProfile table from #TempRoom3
DECLARE temp_cursor6 CURSOR FOR
SELECT iRoomId, iExamineeId FROM #TempRoom3
OPEN temp_cursor6
FETCH NEXT FROM temp_cursor6 INTO @iRoomId, @iExamineeId
WHILE @@FETCH_STATUS = 0 
BEGIN
	SELECT @Id = (SELECT iRoomProfileId FROM #RoomDetail WHERE iRoomId = @iRoomId)	
	SELECT @dtExamDay3 = (SELECT dtDay3 FROM #TempSecondExam)
	-- update dtSecondExamDay field of tbSTEExamineeProfile for the particular examinee
	UPDATE tbSTEExamineeProfile SET dtSecondExamDay = @dtExamDay3 
	WHERE iExamineeProfileId = @iExamineeId
	
	-- get the subjectid for the normal interview
	SELECT @iNormalInterview = (SELECT iSubjectProfileId FROM tbSTESubjectProfile WHERE iExamType = 2)
	-- delete any existing row from tbSTEExamineeRoomProfile for the selected examinee-subject combination
	DELETE FROM tbSTEExamineeRoomProfile 
	WHERE iExamineeProfileId = @iExamineeId
	AND iSubjectProfileId = @iNormalInterview
	-- get the new id for the tbSTEExamineeRoomProfile table
	SELECT @iNewId = (SELECT MAX(iExamineeRoomProfileId) FROM tbSTEExamineeRoomProfile)
	IF @iNewId IS NULL
		SELECT @iNewId = (SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName = 'tbSTEExamineeRoomProfile')
	ELSE
		SET @iNewId = @iNewId + 1
		
		-- insert the examinee data for nomal interview into tbSTEEXamineeRoomProfile for the selected examinee
		INSERT INTO tbSTEExamineeRoomProfile VALUES(@iNewId, @iExamineeId, @Id, @iNormalInterview, getdate(),getdate())
		
	FETCH NEXT FROM temp_cursor6 INTO @iRoomId, @iExamineeId
END
CLOSE temp_cursor6
DEALLOCATE temp_cursor6
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

-- ****************************************************************************************************************************
-- Stored Procedure		: UspCTMAllocateRoom
-- Input Parametrs		: Nil
-- Created					: 22/10/2001		
-- Author					: Dileep Cherian
-- Output					: Recordsets
-- Modification History	: 
-- Reference				: Functional Spec of Distribution Of Examinee.doc (ver 1.0)
-- ****************************************************************************************************************************
CREATE PROCEDURE UspCTMAllocateRoom
AS
BEGIN
SET NOCOUNT ON
DECLARE
@iExamineeId INTEGER,
@iSex INTEGER,
@iDay1Flag INTEGER,
@iDay2Flag INTEGER,
@iDay3Flag INTEGER,
@iMultipleFlag INTEGER,
@iHighSchoolId INTEGER,
@iNoOfExamineeDay1 INTEGER,
@iNoOfExamineeDay2 INTEGER,
@iNoOfExamineeDay3 INTEGER,
@iCountDay1 INTEGER,
@iCountDay2 INTEGER,
@iCountDay3 INTEGER,
@iNoOfConditions INTEGER,
@MultipleValue INTEGER;
/* Store tbSTESecondExamProfile table into this tempporary table */
CREATE TABLE #TempSecondExam
(
	dtDay1 DATETIME,
	dtDay2 DATETIME,
	dtDay3 DATETIME,
	iNoExamineeDay1 INTEGER,
	iNoExamineeDay2 INTEGER,
	iNoExamineeDay3 INTEGER,
	iNoRoomDay1 INTEGER,
	iNoRoomDay2 INTEGER,
	iNoRoomDay3 INTEGER
)
INSERT INTO #TempSecondExam 
SELECT dtSecondExamDay1, dtSecondExamDay2, dtSecondExamDay3,
iNumberOfExamineeDay1, iNumberOfExamineeDay2, iNumberOfExamineeDay3,
iNumberOfRoomDay1, iNumberOfRoomDay2, iNumberOfRoomDay3 FROM tbSTESecondExamProfile 
WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile WHERE iActiveFlag=1)
SELECT @iNoOfExamineeDay1 = (SELECT iNoExamineeDay1 FROM #TempSecondExam)
SELECT @iNoOfExamineeDay2 = (SELECT iNoExamineeDay2 FROM #TempSecondExam)
SELECT @iNoOfExamineeDay3 = (SELECT iNoExamineeDay3 FROM #TempSecondExam)
CREATE TABLE #TempDay1
(
	iExamineeProfileId INTEGER,	
	iSex INTEGER,
	iPreferenceDay1Flag INTEGER,
	iPreferenceDay2Flag INTEGER,
	iPreferenceDay3Flag INTEGER,
	iMultipleApplyFlag INTEGER,
	iHighSchoolId INTEGER
)
CREATE TABLE #TempDay2
(
	iExamineeProfileId INTEGER,	
	iSex INTEGER,
	iPreferenceDay1Flag INTEGER,
	iPreferenceDay2Flag INTEGER,
	iPreferenceDay3Flag INTEGER,
	iMultipleApplyFlag INTEGER,
	iHighSchoolId INTEGER
)
CREATE TABLE #TempDay3
(
	iExamineeProfileId INTEGER,	
	iSex INTEGER,
	iPreferenceDay1Flag INTEGER,
	iPreferenceDay2Flag INTEGER,
	iPreferenceDay3Flag INTEGER,
	iMultipleApplyFlag INTEGER,
	iHighSchoolId INTEGER
)

insert into #TempDay1 select iExamineeProfileId , iSex , iPreferenceDay1Flag , iPreferenceDay2Flag , iPreferenceDay3Flag , iMultipleApplyFlag , iHighSchoolId from tbSTEExamineeprofile where dtSecondExamDay =  (SELECT dtDay1 FROM #TempSecondExam)
insert into #TempDay2 select iExamineeProfileId , iSex , iPreferenceDay1Flag , iPreferenceDay2Flag , iPreferenceDay3Flag , iMultipleApplyFlag , iHighSchoolId from tbSTEExamineeprofile where dtSecondExamDay =  (SELECT dtDay2 FROM #TempSecondExam)
insert into #TempDay3 select iExamineeProfileId , iSex , iPreferenceDay1Flag , iPreferenceDay2Flag , iPreferenceDay3Flag , iMultipleApplyFlag , iHighSchoolId from tbSTEExamineeprofile where dtSecondExamDay =  (SELECT dtDay3 FROM #TempSecondExam)

exec UspCTMAllocateDay1Room
exec UspCTMAllocateDay2Room
exec UspCTMAllocateDay3Room

END	/* Final End */
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

-- ****************************************************************************************************************************
-- Stored Procedure		: UspCTMAllocateDay1Room
-- Input Parametrs		: Nil
-- Created					: 22/10/2001		
-- Author					: Dileep Cherian
--	Output					: Nil
--	Modification History	: Nil
-- Reference				: Functional Spec of Distribution Of Examinee.doc (ver 1.0)
-- ****************************************************************************************************************************
CREATE PROCEDURE UspCTMAllocateRoomDay1
AS
BEGIN
DECLARE
@iExamineeId INTEGER,			
@iSex INTEGER,
@iHighSchoolId INTEGER,
@iTotalExamineeDay1 INTEGER,		-- total examinees allocated for day1
@iTotalRoomsDay1 INTEGER,			-- total rooms allocated for day1
@Capacity INTEGER,					-- room capacity
@Id INTEGER,							-- examinee id
@counter INTEGER,
@SchoolId INTEGER,
@iRoomId INTEGER,
@dtExamDay1 DATETIME,				-- actual date of day1 exam
@iCapacity INTEGER,
@iCount INTEGER,
@iCounter INTEGER,
@iNewId INTEGER,
@iNormalInterview INTEGER,			
@iTotalRoomsAvailable INTEGER;	-- total available rooms
-- store the details of all available rooms in temporary table
CREATE TABLE #RoomDetail			
(
	iRoomId INTEGER,
	iRoomProfileId INTEGER,
	iMaxCapacity INTEGER
)
-- temporary allocation table
CREATE TABLE #TempRoom1
(
	iRoomId INTEGER,
	iExamineeId INTEGER,
	iSex INTEGER,	
	iHighSchoolId INTEGER
)
-- get these values from UspCTMAllocateRoom
SET @iTotalExamineeDay1 = (SELECT COUNT(*) FROM #TempDay1)
SET @iTotalRoomsDay1 = (SELECT iNoRoomDay1 FROM #TempSecondExam)
SET @counter = 1
-- populate the #RoomDetail table
DECLARE temp_cursor2 CURSOR FOR
-- select only those rooms which are flagged for Interviews
SELECT iRoomProfileId, iMaxCapacity FROM tbSTERoomProfile 
WHERE iMaxCapacity > 0 AND iInterviewRoomFlag = 0 ORDER BY iRoomProfileId
OPEN temp_cursor2
FETCH NEXT FROM temp_cursor2 INTO @Id, @Capacity
WHILE @@FETCH_STATUS = 0 AND @counter <= @iTotalRoomsDay1
BEGIN
	INSERT INTO #RoomDetail VALUES(@counter, @Id, @Capacity)
	SET @counter = @counter + 1
	FETCH NEXT FROM temp_cursor2 INTO @Id, @Capacity
END
CLOSE temp_cursor2
DEALLOCATE temp_cursor2
-- get the number of available rooms
SELECT @iTotalRoomsAvailable = (SELECT COUNT(*) FROM #RoomDetail)
SET @iRoomId = 1
-- loop through all the distinct schools ids - outer loop
DECLARE temp_cursor3 CURSOR FOR
SELECT DISTINCT iHighSchoolId FROM #TempDay1
OPEN temp_cursor3
FETCH NEXT FROM temp_cursor3 INTO @SchoolId
WHILE @@FETCH_STATUS = 0 
BEGIN		
	-- loop through all the male examinees in that school - inner loop (male)
	DECLARE temp_cursor4 CURSOR FOR
	SELECT iExamineeProfileId, iSex, iHighSchoolId FROM #TempDay1 WHERE iSex = 0 AND iHighSchoolId = @SchoolId
	OPEN temp_cursor4
	FETCH NEXT FROM temp_cursor4 INTO @iExamineeId, @iSex, @iHighSchoolId
	WHILE @@FETCH_STATUS = 0 
	BEGIN		
		SET @iCounter = 1
		WHILE @iCounter <= @iTotalRoomsDay1  -- do this till the number of rooms allocated for the day is not exceeded
		BEGIN
			SELECT @iCount = (SELECT COUNT(*) FROM #TempRoom1 WHERE iRoomId = @iRoomId)
			SELECT @iCapacity = (SELECT iMaxCapacity FROM #RoomDetail WHERE iRoomId = @iRoomId)
			
			IF @iCount < @iCapacity	-- check for the capacity of the room
			BEGIN	
				INSERT INTO #TempRoom1	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay1
					SET @iRoomId = 1
				BREAK;
			END
			ELSE	-- capacity of room room is full, so move to the enxt room
			BEGIN
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay1
					-- move to the first room, after reaching the last room
					SET @iRoomId = 1
				INSERT INTO #TempRoom1	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay1
					SET @iRoomId = 1
				BREAK;
			END
			SET @iCounter = @iCounter + 1
		END
		FETCH NEXT FROM temp_cursor4 INTO @iExamineeId, @iSex, @iHighSchoolId
	END
	CLOSE temp_cursor4
	DEALLOCATE temp_cursor4 
	-- loop through all the female examinees in that school - inner loop(female)
	DECLARE temp_cursor5 CURSOR FOR
	SELECT iExamineeProfileId, iSex, iHighSchoolId FROM #TempDay1 WHERE iSex = 1 AND iHighSchoolId = @SchoolId
	OPEN temp_cursor5
	FETCH NEXT FROM temp_cursor5 INTO @iExamineeId, @iSex, @iHighSchoolId
	WHILE @@FETCH_STATUS = 0 
	BEGIN
		SET @iCounter = 1
		WHILE @iCounter <= @iTotalRoomsDay1
		BEGIN
			SELECT @iCount = (SELECT COUNT(*) FROM #TempRoom1 WHERE iRoomId = @iRoomId)
			SELECT @iCapacity = (SELECT iMaxCapacity FROM #RoomDetail WHERE iRoomId = @iRoomId)
			IF @iCount < @iCapacity		-- check for the capacity of the room
			BEGIN	
				INSERT INTO #TempRoom1	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay1
					SET @iRoomId = 1
				BREAK;
			END
			ELSE		-- capacity of room room is full, so move to the next room
			BEGIN
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay1
					-- move to the first room, after reaching the last room
					SET @iRoomId = 1
				INSERT INTO #TempRoom1	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay1
					SET @iRoomId = 1
				BREAK;
			END
			SET @iCounter = @iCounter + 1
		END
		FETCH NEXT FROM temp_cursor5 INTO @iExamineeId, @iSex, @iHighSchoolId
	END
	CLOSE temp_cursor5
	DEALLOCATE temp_cursor5 
	FETCH NEXT FROM temp_cursor3 INTO @SchoolId
END
CLOSE temp_cursor3
DEALLOCATE temp_cursor3
-- actual insertion into tbSTEExamineeProfile and tbSTEExamineeRoomProfile table from #TempRoom1
DECLARE temp_cursor6 CURSOR FOR
SELECT iRoomId, iExamineeId FROM #TempRoom1
OPEN temp_cursor6
FETCH NEXT FROM temp_cursor6 INTO @iRoomId, @iExamineeId
WHILE @@FETCH_STATUS = 0 
BEGIN
	
	SELECT @Id = (SELECT iRoomProfileId FROM #RoomDetail WHERE iRoomId = @iRoomId)	-- get the actual roomprofileid
	SELECT @dtExamDay1 = (SELECT dtDay1 FROM #TempSecondExam)								-- get the exam date for day1
	
	-- update dtSecondExamDay field of tbSTEExamineeProfile for the particular examinee
	UPDATE tbSTEExamineeProfile SET dtSecondExamDay = @dtExamDay1 
	WHERE iExamineeProfileId = @iExamineeId
	
	-- get the subjectid for the normal interview
	SELECT @iNormalInterview = (SELECT iSubjectProfileId FROM tbSTESubjectProfile WHERE iExamType = 2)
	-- delete any existing row from tbSTEExamineeRoomProfile for the selected examinee-subject combination
	DELETE FROM tbSTEExamineeRoomProfile 
	WHERE iExamineeProfileId = @iExamineeId
	AND iSubjectProfileId = @iNormalInterview 
	
	-- get the new id for the tbSTEExamineeRoomProfile table
	SELECT @iNewId = (SELECT MAX(iExamineeRoomProfileId) FROM tbSTEExamineeRoomProfile)
	IF @iNewId IS NULL
		SELECT @iNewId = (SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName = 'tbSTEExamineeRoomProfile')		
	ELSE
		SET @iNewId = @iNewId + 1
	
	-- insert the examinee data for nomal interview into tbSTEEXamineeRoomProfile for the selected examinee
	INSERT INTO tbSTEExamineeRoomProfile VALUES(@iNewId, @iExamineeId, @Id, @iNormalInterview, getdate(),getdate())
		
	FETCH NEXT FROM temp_cursor6 INTO @iRoomId, @iExamineeId
END
CLOSE temp_cursor6
DEALLOCATE temp_cursor6
END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

-- ****************************************************************************************************************************
-- Stored Procedure		: UspCTMAllocateDay2Room
-- Input Parametrs		: Nil
-- Created			: 22/10/2001		
-- Author			: Dileep Cherian
--Output				: Nil
--Modification History		: Nil
-- Reference			: Functional Spec of Distribution Of Examinee.doc (ver 1.0)
-- ****************************************************************************************************************************
CREATE PROCEDURE UspCTMAllocateRoomDay2
AS
BEGIN
DECLARE
@iExamineeId INTEGER,
@iSex INTEGER,
@iHighSchoolId INTEGER,
@iTotalExamineeDay2 INTEGER,		-- total examinees allocated for day2
@iTotalRoomsDay2 INTEGER,			-- total rooms allocated for day1
@Capacity INTEGER,					-- room capacity
@Id INTEGER,							-- examinee id
@counter INTEGER,
@SchoolId INTEGER,
@iRoomId INTEGER,
@dtExamDay2 DATETIME,				-- actual date of day1 exam
@iCapacity INTEGER,
@iCount INTEGER,
@iCounter INTEGER,
@iNewId INTEGER,
@iNormalInterview INTEGER,
@iTotalRoomsAvailable INTEGER;	-- total available rooms
-- store the details of all available rooms in temporary table
CREATE TABLE #RoomDetail
(
	iRoomId INTEGER,
	iRoomProfileId INTEGER,
	iMaxCapacity INTEGER
)
-- temporary allocation table
CREATE TABLE #TempRoom2
(
	iRoomId INTEGER,
	iExamineeId INTEGER,
	iSex INTEGER,	
	iHighSchoolId INTEGER
)
-- get these values from UspCTMAllocateRoom
SET @iTotalExamineeDay2 = (SELECT COUNT(*) FROM #TempDay2)
SET @iTotalRoomsDay2 = (SELECT iNoRoomDay2 FROM #TempSecondExam)
SET @counter = 1
-- populate the #RoomDetail table
DECLARE temp_cursor2 CURSOR FOR
-- select only those rooms which are flagged for Interviews
SELECT iRoomProfileId, iMaxCapacity FROM tbSTERoomProfile 
WHERE iMaxCapacity > 0 AND iInterviewRoomFlag = 0 ORDER BY iRoomProfileId
OPEN temp_cursor2
FETCH NEXT FROM temp_cursor2 INTO @Id, @Capacity
WHILE @@FETCH_STATUS = 0 AND @counter <= @iTotalRoomsDay2
BEGIN
	INSERT INTO #RoomDetail VALUES(@counter, @Id, @Capacity)
	SET @counter = @counter + 1
	FETCH NEXT FROM temp_cursor2 INTO @Id, @Capacity
END
CLOSE temp_cursor2
DEALLOCATE temp_cursor2
-- get the number of available rooms
SELECT @iTotalRoomsAvailable = (SELECT COUNT(*) FROM #RoomDetail)
SET @iRoomId = 1
-- loop through all the distinct schools ids - outer loop
DECLARE temp_cursor3 CURSOR FOR
SELECT DISTINCT iHighSchoolId FROM #TempDay2
OPEN temp_cursor3
FETCH NEXT FROM temp_cursor3 INTO @SchoolId
WHILE @@FETCH_STATUS = 0 
BEGIN
	-- loop through all the male examinees in that school - inner loop (male)
	DECLARE temp_cursor4 CURSOR FOR
	SELECT iExamineeProfileId, iSex, iHighSchoolId FROM #TempDay2 WHERE iSex = 0 AND iHighSchoolId = @SchoolId
	OPEN temp_cursor4
	FETCH NEXT FROM temp_cursor4 INTO @iExamineeId, @iSex, @iHighSchoolId
	WHILE @@FETCH_STATUS = 0 
	BEGIN
		SET @iCounter = 1
		WHILE @iCounter <= @iTotalRoomsDay2		-- do this till the number of rooms allocated for the day is not exceeded
		BEGIN
			SELECT @iCount = (SELECT COUNT(*) FROM #TempRoom2 WHERE iRoomId = @iRoomId)
			SELECT @iCapacity = (SELECT iMaxCapacity FROM #RoomDetail WHERE iRoomId = @iRoomId)
			
IF @iCount < @iCapacity		-- check for the capacity of the room
			BEGIN	
				INSERT INTO #TempRoom2	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay2
					SET @iRoomId = 1
				BREAK;
			END
			ELSE		-- capacity of room room is full, so move to the enxt room
			BEGIN
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay2
					-- move to the first room, after reaching the last room
					SET @iRoomId = 1
				INSERT INTO #TempRoom2	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay2
					SET @iRoomId = 1
				BREAK;
			END
			SET @iCounter = @iCounter + 1
		END
		FETCH NEXT FROM temp_cursor4 INTO @iExamineeId, @iSex, @iHighSchoolId
	END
	CLOSE temp_cursor4
	DEALLOCATE temp_cursor4 
	
	-- loop through all the female examinees in that school - inner loop(female)
	DECLARE temp_cursor5 CURSOR FOR
	SELECT iExamineeProfileId, iSex, iHighSchoolId FROM #TempDay2 WHERE iSex = 1 AND iHighSchoolId = @SchoolId
	OPEN temp_cursor5
	FETCH NEXT FROM temp_cursor5 INTO @iExamineeId, @iSex, @iHighSchoolId
	WHILE @@FETCH_STATUS = 0 
	BEGIN
		SET @iCounter = 1
		WHILE @iCounter <= @iTotalRoomsDay2
		BEGIN
			SELECT @iCount = (SELECT COUNT(*) FROM #TempRoom2 WHERE iRoomId = @iRoomId)
			SELECT @iCapacity = (SELECT iMaxCapacity FROM #RoomDetail WHERE iRoomId = @iRoomId)
			IF @iCount < @iCapacity		-- check for the capacity of the room
			BEGIN	
				INSERT INTO #TempRoom2	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay2
					SET @iRoomId = 1
				BREAK;
			END
			ELSE		-- capacity of room room is full, so move to the next room
			BEGIN
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay2
					-- move to the first room, after reaching the last room
					SET @iRoomId = 1
				INSERT INTO #TempRoom2	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay2
					SET @iRoomId = 1
				BREAK;
			END
			SET @iCounter = @iCounter + 1
		END
		FETCH NEXT FROM temp_cursor5 INTO @iExamineeId, @iSex, @iHighSchoolId
	END
	CLOSE temp_cursor5
	DEALLOCATE temp_cursor5 
	FETCH NEXT FROM temp_cursor3 INTO @SchoolId
END
CLOSE temp_cursor3
DEALLOCATE temp_cursor3
-- actual insertion into tbSTEExamineeProfile and tbSTEExamineeRoomProfile table from #TempRoom2
DECLARE temp_cursor6 CURSOR FOR
SELECT iRoomId, iExamineeId, iSex, iHighSchoolId FROM #TempRoom2
OPEN temp_cursor6
FETCH NEXT FROM temp_cursor6 INTO @iRoomId, @iExamineeId, @iSex, @iHighSchoolId
WHILE @@FETCH_STATUS = 0 
BEGIN
	SELECT @Id = (SELECT iRoomProfileId FROM #RoomDetail WHERE iRoomId = @iRoomId)
	SELECT @dtExamDay2 = (SELECT dtDay2 FROM #TempSecondExam)
	
	-- update dtSecondExamDay field of tbSTEExamineeProfile for the particular examinee
	UPDATE tbSTEExamineeProfile SET dtSecondExamDay = @dtExamDay2 
	WHERE iExamineeProfileId = @iExamineeId
	
	-- get the subjectid for the normal interview
	SELECT @iNormalInterview = (SELECT iSubjectProfileId FROM tbSTESubjectProfile WHERE iExamType = 2)
	-- delete any existing row from tbSTEExamineeRoomProfile for the selected examinee-subject combination
	DELETE FROM tbSTEExamineeRoomProfile 
	WHERE iExamineeProfileId = @iExamineeId
	AND iSubjectProfileId = @iNormalInterview 
	-- get the new id for the tbSTEExamineeRoomProfile table
	SELECT @iNewId = (SELECT MAX(iExamineeRoomProfileId) FROM tbSTEExamineeRoomProfile)
	IF @iNewId IS NULL
		SELECT @iNewId = (SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName = 'tbSTEExamineeRoomProfile')
	ELSE
		SET @iNewId = @iNewId + 1
	
	-- insert the examinee data for nomal interview into tbSTEEXamineeRoomProfile for the selected examinee
	INSERT INTO tbSTEExamineeRoomProfile VALUES(@iNewId, @iExamineeId, @Id, @iNormalInterview, getdate(),getdate())
		
	FETCH NEXT FROM temp_cursor6 INTO @iRoomId, @iExamineeId, @iSex, @iHighSchoolId
END
CLOSE temp_cursor6
DEALLOCATE temp_cursor6
END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

-- ****************************************************************************************************************************
-- Stored Procedure		: UspCTMAllocateDay3Room
-- Input Parametrs		: Nil
-- Created					: 22/10/2001		
-- Author					: Dileep Cherian
--Output						: Nil
--Modification History	: Nil
-- Reference				: Functional Spec of Distribution Of Examinee.doc (ver 1.0)
-- ****************************************************************************************************************************
CREATE PROCEDURE UspCTMAllocateRoomDay3
AS
BEGIN
DECLARE
@iExamineeId INTEGER,
@iSex INTEGER,
@iHighSchoolId INTEGER,
@iTotalExamineeDay3 INTEGER,		-- total examinees allocated for day1
@iTotalRoomsDay3 INTEGER,			-- total rooms allocated for day1
@Capacity INTEGER,					-- room capacity
@Id INTEGER,							-- examinee id
@counter INTEGER,
@SchoolId INTEGER,
@iRoomId INTEGER,
@dtExamDay3 DATETIME,				-- actual date of day3 exam
@iCapacity INTEGER,
@iCount INTEGER,
@iCounter INTEGER,
@iNewId INTEGER,
@iNormalInterview INTEGER,
@iTotalRoomsAvailable INTEGER;	-- total available rooms
-- store the details of all available rooms in temporary table
CREATE TABLE #RoomDetail
(
	iRoomId INTEGER,
	iRoomProfileId INTEGER,
	iMaxCapacity INTEGER
)
-- temporary allocation table
CREATE TABLE #TempRoom3
(
	iRoomId INTEGER,
	iExamineeId INTEGER,
	iSex INTEGER,	
	iHighSchoolId INTEGER
)
-- get these values from UspCTMAllocateRoom
SET @iTotalExamineeDay3 = (SELECT COUNT(*) FROM #TempDay3)
SET @iTotalRoomsDay3 = (SELECT iNoRoomDay3 FROM #TempSecondExam)
SET @counter = 1
-- populate the #RoomDetail table
DECLARE temp_cursor2 CURSOR FOR
-- select only those rooms which are flagged for Interviews
SELECT iRoomProfileId, iMaxCapacity FROM tbSTERoomProfile 
WHERE iMaxCapacity > 0 AND iInterviewRoomFlag = 0 ORDER BY iRoomProfileId
OPEN temp_cursor2
FETCH NEXT FROM temp_cursor2 INTO @Id, @Capacity
WHILE @@FETCH_STATUS = 0 AND @counter <= @iTotalRoomsDay3
BEGIN
	INSERT INTO #RoomDetail VALUES(@counter, @Id, @Capacity)
	SET @counter = @counter + 1
	FETCH NEXT FROM temp_cursor2 INTO @Id, @Capacity
END
CLOSE temp_cursor2
DEALLOCATE temp_cursor2
-- get the number of available rooms
SELECT @iTotalRoomsAvailable = (SELECT COUNT(*) FROM #RoomDetail)
SET @iRoomId = 1
-- loop through all the distinct schools ids - outer loop
DECLARE temp_cursor3 CURSOR FOR
SELECT DISTINCT iHighSchoolId FROM #TempDay3
OPEN temp_cursor3
FETCH NEXT FROM temp_cursor3 INTO @SchoolId
WHILE @@FETCH_STATUS = 0 
BEGIN
	-- loop through all the male examinees in that school - inner loop (male)
	DECLARE temp_cursor4 CURSOR FOR
	SELECT iExamineeProfileId, iSex, iHighSchoolId FROM #TempDay3 WHERE iSex = 0 AND iHighSchoolId = @SchoolId
	OPEN temp_cursor4
	FETCH NEXT FROM temp_cursor4 INTO @iExamineeId, @iSex, @iHighSchoolId
	WHILE @@FETCH_STATUS = 0 
	BEGIN
		SET @iCounter = 1
		WHILE @iCounter <= @iTotalRoomsDay3		-- do this till the number of rooms allocated for the day is not exceeded
		BEGIN
			SELECT @iCount = (SELECT COUNT(*) FROM #TempRoom3 WHERE iRoomId = @iRoomId)
			SELECT @iCapacity = (SELECT iMaxCapacity FROM #RoomDetail WHERE iRoomId = @iRoomId)
			
			IF @iCount < @iCapacity		-- check for the capacity of the room
			BEGIN	
				INSERT INTO #TempRoom3	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay3
					SET @iRoomId = 1
				BREAK;
			END
			ELSE		-- capacity of room room is full, so move to the enxt room
			BEGIN
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay3
					-- move to the first room, after reaching the last room
					SET @iRoomId = 1
				INSERT INTO #TempRoom3	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay3
					SET @iRoomId = 1
				BREAK;
			END
			SET @iCounter = @iCounter + 1
		END
		FETCH NEXT FROM temp_cursor4 INTO @iExamineeId, @iSex, @iHighSchoolId
	END
	CLOSE temp_cursor4
	DEALLOCATE temp_cursor4 
	-- loop through all the female examinees in that school - inner loop(female)
	DECLARE temp_cursor5 CURSOR FOR
	SELECT iExamineeProfileId, iSex, iHighSchoolId FROM #TempDay3 WHERE iSex = 1 AND iHighSchoolId = @SchoolId
	OPEN temp_cursor5
	FETCH NEXT FROM temp_cursor5 INTO @iExamineeId, @iSex, @iHighSchoolId
	WHILE @@FETCH_STATUS = 0 
	BEGIN
		SET @iCounter = 1
		WHILE @iCounter <= @iTotalRoomsDay3
		BEGIN
			SELECT @iCount = (SELECT COUNT(*) FROM #TempRoom3 WHERE iRoomId = @iRoomId)
			SELECT @iCapacity = (SELECT iMaxCapacity FROM #RoomDetail WHERE iRoomId = @iRoomId)
			IF @iCount < @iCapacity		-- check for the capacity of the room
			BEGIN	
				INSERT INTO #TempRoom3	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay3
					SET @iRoomId = 1
				BREAK;
			END
			ELSE		-- capacity of room room is full, so move to the next room
			BEGIN
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay3
					-- move to the first room, after reaching the last room
					SET @iRoomId = 1
				INSERT INTO #TempRoom3	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsAvailable OR @iRoomId > @iTotalRoomsDay3
					SET @iRoomId = 1
				BREAK;
			END
			SET @iCounter = @iCounter + 1
		END
		--PRINT @iRoomId
		FETCH NEXT FROM temp_cursor5 INTO @iExamineeId, @iSex, @iHighSchoolId
	END
	CLOSE temp_cursor5
	DEALLOCATE temp_cursor5 
	FETCH NEXT FROM temp_cursor3 INTO @SchoolId
END
CLOSE temp_cursor3
DEALLOCATE temp_cursor3
-- actual insertion into tbSTEExamineeProfile and tbSTEExamineeRoomProfile table from #TempRoom3
DECLARE temp_cursor6 CURSOR FOR
SELECT iRoomId, iExamineeId FROM #TempRoom3
OPEN temp_cursor6
FETCH NEXT FROM temp_cursor6 INTO @iRoomId, @iExamineeId
WHILE @@FETCH_STATUS = 0 
BEGIN
	SELECT @Id = (SELECT iRoomProfileId FROM #RoomDetail WHERE iRoomId = @iRoomId)	
	SELECT @dtExamDay3 = (SELECT dtDay3 FROM #TempSecondExam)
	-- update dtSecondExamDay field of tbSTEExamineeProfile for the particular examinee
	UPDATE tbSTEExamineeProfile SET dtSecondExamDay = @dtExamDay3 
	WHERE iExamineeProfileId = @iExamineeId
	
	-- get the subjectid for the normal interview
	SELECT @iNormalInterview = (SELECT iSubjectProfileId FROM tbSTESubjectProfile WHERE iExamType = 2)
	-- delete any existing row from tbSTEExamineeRoomProfile for the selected examinee-subject combination
	DELETE FROM tbSTEExamineeRoomProfile 
	WHERE iExamineeProfileId = @iExamineeId
	AND iSubjectProfileId = @iNormalInterview
	-- get the new id for the tbSTEExamineeRoomProfile table
	SELECT @iNewId = (SELECT MAX(iExamineeRoomProfileId) FROM tbSTEExamineeRoomProfile)
	IF @iNewId IS NULL
		SELECT @iNewId = (SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName = 'tbSTEExamineeRoomProfile')
	ELSE
		SET @iNewId = @iNewId + 1
		
		-- insert the examinee data for nomal interview into tbSTEEXamineeRoomProfile for the selected examinee
		INSERT INTO tbSTEExamineeRoomProfile VALUES(@iNewId, @iExamineeId, @Id, @iNormalInterview, getdate(),getdate())
		
	FETCH NEXT FROM temp_cursor6 INTO @iRoomId, @iExamineeId
END
CLOSE temp_cursor6
DEALLOCATE temp_cursor6
END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

-- ****************************************************************************************************************************
-- Stored Procedure		: UspCTMCalScore
-- Input Parametrs		: Nil
-- Created			: 14/09/2001		
-- Author			: Dileep Cherian
--Output				: float
--Modification History		: Nil
-- Reference			: Functional Spec of Distribution Of Examinee.doc (ver 1.0)
-- ****************************************************************************************************************************
CREATE  PROCEDURE UspCTMCalScore (
@ExamType int,
@SubjectProfileId int,
@NumberOfParams int,
@Score1 int,
@Score2 int,
@Score3 int,
@Score4 int,
@Score5 int,
@Score6 int,
@Score7 int,
@Score8 int,
@Score9 int,
@Score10 int,
@TotalScore decimal OUTPUT
)
--RETURNS int
AS
BEGIN
	--DECLARE @TotalScore decimal
	
	IF @NumberOfParams = 1
		SET @TotalScore = @Score1
	
	IF @NumberOfParams = 2
		SET @TotalScore = (@Score1+@Score2)/2
	
	IF @NumberOfParams = 3
		SET @TotalScore = (@Score1+@Score2+@Score3)/3
	
	IF @NumberOfParams = 4
		SET @TotalScore = (@Score1+@Score2+@Score3+@Score4)/4
	
	IF @NumberOfParams = 5
		SET @TotalScore = (@Score1+@Score2+@Score3+@Score4+@Score5)/5
	
	IF @NumberOfParams = 6
		SET @TotalScore = (@Score1+@Score2+@Score3+@Score4+@Score5+@Score6)/6
	
	IF @NumberOfParams = 7
		SET @TotalScore = (@Score1+@Score2+@Score3+@Score4+@Score5+@Score6+@Score7)/7
	
	IF @NumberOfParams = 8
		SET @TotalScore = (@Score1+@Score2+@Score3+@Score4+@Score5+@Score6+@Score7+@Score8)/8		
	
	IF @NumberOfParams = 9
		SET @TotalScore = (@Score1+@Score2+@Score3+@Score4+@Score5+@Score6+@Score7+@Score8+@Score9)/9
		
	IF @NumberOfParams = 10
		SET @TotalScore = (@Score1+@Score2+@Score3+@Score4+@Score5+@Score6+@Score7+@Score8+@Score9+@Score10)/10
	
	RETURN(@TotalScore)
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE     PROCEDURE uspSTEConvertExaminee
						@iNendo int , -- Nendo
						@iJukenNumber	int ,
						@vExamineeName  	varchar(50) = NULL ,
						@vKanaName 		varchar(50) = NULL ,
						@dtBirthDay 		datetime ,
						@vSex 			varchar(1)  ,
						@vHcode 		varchar(6) ,
						@iCourse int = 0 ,
						@iDepartment int = 0 ,
						@iAdmissionType1 int,		--	%genrou1% ,
						@iBackgroundId	int =0, 
						@iFamilyId	int =0,			--	%family% ,
						@iParentJobCategory	int =0 ,	-- %job% ,
						@iQualificationId	int =0	,	-- %sikaku% ,
						@iSuisenFlagId	int = 0	,	-- %suisen% ,
						@vUnivName		varchar(50) ,	--	'%univname%' ,
						@iUniversityType	int = 0	,	--	%univtype% ,
						@vNationality		varchar(50) ,	--	'%kuni%' ,
						@iPhysicalConditionId	int = 0 ,	--	%kenkou% ,
						@iLanguageSubject	int , -- 選択外国語
						@iScienceSub1	int , -- 選択理科
						@iScienceSub2	int , -- 選択理科
						@vHyoteiGrade	varchar(2),	--	'%hyti_seiseki%' ,
						@vRejectFlag		varchar(2) ,	--	'%stat.f1%' ,
						@iExamineeStatus	int = 0 ,	--	%stat.f2% ,
						@iPreferenceDay1Flag	int = 0 , -- 面接希望日１
						@iPreferenceDay2Flag	int = 0 , -- 面接希望日2
						@iPreferenceDay3Flag	int = 0 , -- 面接希望日3
						@iMultipleApplyFlag	int = 0  ,
						@vZipcode 		varchar(7) ,
						@vAddress varchar(255) ,
						@iAbsentFlag int = 0 
AS
declare	@iExamineeProfileID int,
	@iRejectFlag int ,
	@iSex int ,
	@iAdmissionType2 int ,
	@iZipCodeId int,
	@iHighSchoolId int ,
	@vZipAddress varchar(255),
	@vPatAddress varchar(255) ,
	@iLanguageSubjectProfileId int ,
	@iScienceSubProfileId1 int ,
	@iScienceSubProfileId2 int 
-- ProfileIDの最大値を求める
 SELECT  @iExamineeProfileID=MAX( iExamineeProfileID)+1 FROM tbSTEExamineeProfile
  SET @iAdmissionType2=@iAdmissionType1
  IF ( @vSex = 'M' )
   SET @iSex=0
 ELSE
   SET @iSex=1
  IF ( @vRejectFlag = '-' )
   SET @iRejectFlag=1
 ELSE
   SET @iRejectFlag=0
  SELECT  @iBackgroundId=@iBackgroundId+100   -- Space
  SELECT  @iUniversityType=@iUniversityType+200   -- Space
  SELECT  @iFamilyId=@iFamilyId+300   -- Space
  SELECT  @iParentJobCategory=@iParentJobCategory+400   -- Space
  SELECT  @iQualificationId=@iQualificationId+500   -- Space
  SELECT  @iPhysicalConditionId=@iPhysicalConditionId+200   -- Space
/*
  SET @iQualificationId=500    -- Space
						@iBackgroundId	int , 
						@iFamilyId	int ,			--	%family% ,
						@iParentJobCategory	int ,	-- %job% ,
						@iQualificationId	int	,	-- %sikaku% ,
						@iPhysicalConditionId	int ,	--	%kenkou% ,
*/
    -- 1:物理（１０）　２：化学（１１）　３：生物（１２）
  IF ( @iScienceSub1 = 1 )
	SET @iScienceSubProfileId1=14
  ELSE IF ( @iScienceSub1 = 2 )
	SET @iScienceSubProfileId1=15
  ELSE IF ( @iScienceSub1 = 3 )
	SET  @iScienceSubProfileId1=16
  IF ( @iScienceSub2 = 1 )
	SELECT @iScienceSubProfileId2=14
  ELSE IF ( @iScienceSub2 = 2 )
	SELECT @iScienceSubProfileId2=15
  ELSE IF ( @iScienceSub2 = 3 )
	SELECT @iScienceSubProfileId2=16
  IF ( @iLanguageSubject = 0 )
	SET @iLanguageSubjectProfileId=10
  ELSE IF ( @iLanguageSubject=1 )
	SET @iLanguageSubjectProfileId=11
  ELSE IF ( @iLanguageSubject=2 )
	SET @iLanguageSubjectProfileId=12
  ELSE
	SET @iLanguageSubjectProfileId=10
-- 高校コードを検索する。
   SET @iHighSchoolId=NULL
    SELECT @iHighSchoolId=iHighSchoolId FROM tbSTEHighSchoolType
	WHERE vHighSchoolCode=@vHcode
    IF( @iHighSchoolId IS  NULL )
    BEGIN
      SELECT @iHighSchoolId=MAX(iHighSchoolId)+1 FROM tbSTEHighSchoolType
      INSERT INTO  tbSTEHighSchoolType ( iHighSchoolId , vHighSchoolCode , dtCreate , dtUpdate)
		VALUES ( @iHighSchoolId , @vHcode , getdate() , getdate() )
    END
    SELECT @vAddress=RTRIM(REPLACE(@vAddress,'　',' '))
    SELECT @vAddress=REPLACE(@vAddress,' ','')
-- 郵便番号辞書を検索する。
--
-- 郵便番号を７桁にする。
--
SELECT @vZipCode=RTRIM(@vZipCode)
  WHILE LEN( @vZipCode ) < 7 
    BEGIN
      SELECT @vZipCode=@vZipCode+'0'
    END
    SET @iZipCodeId=NULL
    SELECT TOP 1 @iZipCodeId=iZipCodeId,@vZipaddress=vAddress1, @vPatAddress='%'+vAddress1+'%' FROM tbSTEZipCodeMaster
	WHERE vZipCodeName=@vZipCode ORDER BY vZipCodeName
    IF( @iZipCodeId IS NOT NULL )
     BEGIN
   --
   -- 自動インサートはしない。
   -- アドレスを最適な長さに切り詰める。（郵便番号辞書以降の文字列に切り詰める）
	IF ( PATINDEX(@vPatAddress,@vAddress) > 0 )
	      SELECT @vAddress=SUBSTRING(@vAddress,(LEN(@vZipaddress)+PATINDEX(@vPatAddress,@vAddress)) , 255)
      END
    ELSE
      BEGIN
	      SELECT @iZipCodeId=CONVERT(int , @vZipCode )
	      INSERT INTO  tbSTEZipCodeMaster ( iZipCodeId ,  vZipCodeName , dtCreate , dtUpdate)
		VALUES ( @iZipCodeId , @vZipCode , getdate() , getdate() )
      END
SELECT @iZipCodeId,@vZipaddress, @vPatAddress,@vZipCode
-- Insert new records into ExamineeTable
-- SET @iBackgroundId=200
insert into tbSTEExamineeProfile
 (
	iExamineeProfileID,
	iJukenNumber ,
	iNendo ,
	vExamineeName ,
	vKanaName ,
	vAddress ,
	iZipCodeId ,
	iSex ,
	iHighSchoolId ,
	dtBirthDay ,
	iUniversityType ,
	iBackgroundId ,
	iQualificationId ,
	iLanguageSubjProfileId ,
	iScienceSubjProfileId1 ,
	iScienceSubjProfileId2 , 
	iPreferenceDay1Flag ,
	iPreferenceDay2Flag ,
	iPreferenceDay3Flag ,
	iMultipleApplyFlag ,
	iAdmissionType1 ,
	iAdmissionType2 ,
	dtCreate ,
	dtUpdate ,
	iExamineeStatus,
	iAbsentFlag ,
	iRejectFlag ,
	iCourse ,
	iDepartment ,
	iFamilyId ,
	iParentJobCategory ,
	iSuisenFlagId ,
	vUnivName ,
	vNationality ,
	iPhysicalConditionId ,
	vHyoteiGrade
 )
 VALUES 
(
	ISNULL(@iExamineeProfileID,1),
	@iJukenNumber ,
	@iNendo ,
	@vExamineeName ,
	@vKanaName ,
	@vAddress ,
	@iZipCodeId ,
	@iSex ,
	@iHighSchoolId ,
	@dtBirthDay ,
	@iUniversityType ,   -- Space
	@iBackgroundId ,
	@iQualificationId ,    -- Space
	@iLanguageSubjectProfileId ,
	@iScienceSubProfileId1 ,
	@iScienceSubProfileId2 , 
	@iPreferenceDay1Flag ,
	@iPreferenceDay2Flag ,
	@iPreferenceDay3Flag ,
	@iMultipleApplyFlag ,
	@iAdmissionType1 ,
	@iAdmissionType2 ,
	getdate() ,
	getdate() ,
	@iExamineeStatus ,
	@iAbsentFlag ,
	@iRejectFlag ,
	@iCourse ,
	@iDepartment ,
	@iFamilyId ,
	@iParentJobCategory ,
	@iSuisenFlagId ,
	@vUnivName ,
	@vNationality ,
	@iPhysicalConditionId ,
	@vHyoteiGrade
)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE  PROCEDURE uspSTEConvertModify	@iNendo	int ,
						@iJukenNumber int ,
						@iStat0	int =0 ,
						@dtModify0	datetime = NULL ,
						@iStat1	int =0 ,
						@dtModify1	datetime = NULL ,
						@iStat2	int =0 ,
						@dtModify2	datetime = NULL ,
						@iStat3	int =0 ,
						@dtModify3	datetime = NULL ,
						@iStat4	int =0 ,
						@dtModify4	datetime = NULL 
AS
 DECLARE @iExamineeProfileId int 
 SELECT @iExamineeProfileId=iExamineeProfileId FROM tbSTEExamineeProfile
	WHERE iNendo=@iNendo AND iJukenNumber=@iJukenNumber
 IF ( @iStat0 <> 0 )
  BEGIN
	INSERT INTO tbSTEExamineeStatusTrail
			(
				iExamineeProfileId ,
				iPos ,
				iExamineeStatus ,
				iRejectFlag ,
				dtModify ,
				dtCreate
			)
			VALUES
			(
				@iExamineeProfileId ,
				0 ,
				ABS(@iStat0) ,
				(
				CASE	 WHEN @iStat0 < 0 THEN 1 ELSE 0 END 
				)  ,
				@dtModify0 ,
				getdate()
			)
  END
 IF ( @iStat1 <> 0 )
  BEGIN
	INSERT INTO tbSTEExamineeStatusTrail
			(
				iExamineeProfileId ,
				iPos ,
				iExamineeStatus ,
				iRejectFlag ,
				dtModify ,
				dtCreate
			)
			VALUES
			(
				@iExamineeProfileId ,
				1 ,
				ABS(@iStat1) ,
				(
				CASE	 WHEN @iStat1 < 0 THEN 1 ELSE 0 END 
				)  ,
				@dtModify1 ,
				getdate()
			)
  END
 IF ( @iStat2 <> 0 )
  BEGIN
	INSERT INTO tbSTEExamineeStatusTrail
			(
				iExamineeProfileId ,
				iPos ,
				iExamineeStatus ,
				iRejectFlag ,
				dtModify ,
				dtCreate
			)
			VALUES
			(
				@iExamineeProfileId ,
				2 ,
				ABS(@iStat2) ,
				(
				CASE	 WHEN @iStat2 < 0 THEN 1 ELSE 0 END 
				)  ,
				@dtModify2 ,
				getdate()
			)
  END
 IF ( @iStat3 <> 0 )
  BEGIN
	INSERT INTO tbSTEExamineeStatusTrail
			(
				iExamineeProfileId ,
				iPos ,
				iExamineeStatus ,
				iRejectFlag ,
				dtModify ,
				dtCreate
			)
			VALUES
			(
				@iExamineeProfileId ,
				3 ,
				ABS(@iStat3) ,
				(
				CASE	 WHEN @iStat3 < 0 THEN 1 ELSE 0 END 
				)  ,
				@dtModify3 ,
				getdate()
			)
  END
 IF ( @iStat4 <> 0 )
  BEGIN
	INSERT INTO tbSTEExamineeStatusTrail
			(
				iExamineeProfileId ,
				iPos ,
				iExamineeStatus ,
				iRejectFlag ,
				dtModify ,
				dtCreate
			)
			VALUES
			(
				@iExamineeProfileId ,
				4 ,
				ABS(@iStat4) ,
				(
				CASE	 WHEN @iStat4 < 0 THEN 1 ELSE 0 END 
				)  ,
				@dtModify4 ,
				getdate()
			)
  END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE  PROCEDURE uspSTEConvertSchool
				@vHighSchoolCode	varchar(6),	-- '%hcode%' ,
				@vHighSchoolName	varchar(255),	-- '%hname%' ,
				@vZipCode		varchar(8),	-- '%post%' ,
				@vAddress1		varchar(255) , -- '%addr1%',
				@vAddress2		varchar(255) ,-- '%addr2%' ,
				@vTelephoneNo	varchar(20),	-- '%tel%' ,
				@vFaxNo		varchar(20),	-- '%fax%' ,
				@vRepresentiveName	varchar(255) ,-- '%ontmei%' ,
				@iLetterFlag		int ,		-- %soufu% ,
				@iHighSchoolRecommendation	int ,		-- %sitei% ,
				@iHighSchoolRecommendationYear1	int ,-- %sitei_nen1% ,
				@iHighSchoolRecommendationYear2	int ,-- %sitei_nen2% ,
				@iHighSchoolDropRecommendationYear	int -- %sitei_kaijo% ,
AS
 DECLARE 	@iHighSchoolId int ,
		@iZipCodeId int
   SET @iHighSchoolId=NULL
    SELECT @iHighSchoolId=iHighSchoolId FROM tbSTEHighSchoolType
	WHERE vHighSchoolCode=@vHighSchoolCode
    IF( @iHighSchoolId IS  NULL )
    BEGIN
      SELECT @iHighSchoolId=MAX(iHighSchoolId)+1 FROM tbSTEHighSchoolType
      --  無いので追加する。
  SELECT @vZipCode=RTRIM(@vZipCode)
  WHILE LEN( @vZipCode ) < 7 
      BEGIN
        SELECT @vZipCode=@vZipCode+'0'
      END
    SET @iZipCodeId=NULL
    SELECT TOP 1 @iZipCodeId=iZipCodeId FROM tbSTEZipCodeMaster
	WHERE vZipCodeName=@vZipCode ORDER BY vZipCodeName
     INSERT INTO tbSTEHIghSchoolType
		(
			iHighSchoolId,
			iZipCodeId ,
			vHighSchoolCode ,
			vHighSchoolName ,
			vAddress1	,
			vAddress2	,
			vTelephoneNo ,
			vFaxNo	,
			vRepresentativeName ,
			iLetterFlag	,
			iHighSchoolRecommendation ,
			iHighSchoolRecommendationYear1 ,
			iHighSchoolRecommendationYear2 ,
			iHighSchoolDropRecommendationYear ,
			dtCreate ,
			dtUpdate
		)
		VALUES
		(
			ISNULL(@iHighSchoolId,1) ,
			@iZipCodeId ,
			@vHighSchoolCode ,
			@vHighSchoolName ,
			@vAddress1	,
			@vAddress2	,
			@vTelephoneNo ,
			@vFaxNo	,
			@vRepresentiveName ,
			@iLetterFlag	,
			@iHighSchoolRecommendation ,
			@iHighSchoolRecommendationYear1 ,
			@iHighSchoolRecommendationYear2 ,
			@iHighSchoolDropRecommendationYear ,
			getdate() ,
			getdate()
		)
    END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE   PROCEDURE uspSTEConvertScore	@iSubjectProfileId	int ,
						@iNendo			int ,
						@iJukenNumber		int ,
						@fRawScore		float ,
						@fChoseiScore		float ,
						@iSelFlag	int ,
						@iAbsentFlag		int
AS
 DECLARE 	@iScoreProfileId	int ,
		@iExamineeProfileId  int ,
		@iScoreDetailId	int ,
		@iSubjectQuestionId int
 SELECT @iExamineeProfileId=iExamineeProfileId FROM tbSTEExamineeProfile
	WHERE iNendo=@iNendo AND iJukenNumber=@iJukenNumber
 -- 
 SELECT @iScoreProfileId=iScoreProfileId FROM tbSTEScoreProfile
	WHERE iSubjectProfileId=@iSubjectProfileId AND iExamineeProfileId=@iExamineeProfileId
-- SELECT @iSelFlag as iSelFlag , @iScoreProfileId as iScoreProfileId, @iExamineeProfileId as iExamineeProfileId
 IF ( (@iSelFlag <> 0) AND (@iScoreProfileId IS NULL)  AND  (@iExamineeProfileId IS NOT NULL) )
  BEGIN
     SELECT @iScoreProfileId=MAX(iScoreProfileId)+1 FROM tbSTEScoreProfile
     SELECT @iScoreDetailId=MAX(iScoreDetailId)+1   FROM tbSTEScoreDetail
-- SELECT @iSelFlag as iSelFlag , @iScoreProfileId as iScoreProfileId, @iExamineeProfileId as iExamineeProfileId
     --
     INSERT INTO tbSTEScoreProfile
		(
			iScoreProfileId ,
			iSubjectProfileId ,
			iExamineeProfileId ,
			fRawScore ,
			fChoseiScore ,
			iAbsentFlag ,
			dtCreate ,
			dtUpdate
		)
		VALUES
		(
			ISNULL( @iScoreProfileId,1) ,
			@iSubjectProfileId ,
			@iExamineeProfileId ,
			@fRawScore ,
			@fChoseiScore ,
			@iAbsentFlag ,
			getdate() ,
			getdate()
		)
     --
     SELECT TOP 1 @iSubjectQuestionId=iSubjectQuestionId   FROM tbSTESubjectQuestionProfile
	WHERE iSubjectProfileId=@iSubjectProfileId
     --
     INSERT INTO tbSTEScoreDetail
		(
			iScoreDetailId ,
			iScoreProfileId ,
			iSubjectQuestionId ,
			fDetailScore ,
			dtCreate ,
			dtUpdate
		)
		VALUES
		(
			ISNULL(@iScoreDetailId,1) ,
			@iScoreProfileId ,
			@iSubjectQuestionId ,
			@fRawScore ,
			getdate() ,
			getdate()
		)
  END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

/*
 *  Insert into tbSTEExamineeProfile from CSV file 
 *
 */
CREATE PROCEDURE uspSTEInsertExaminee	@iNendo int , -- Nendo
						@iJukenNumber	int ,
						@vKanaName 		varchar(50) = NULL ,
						@vExamineeName  	varchar(50) = NULL ,
						@dtBirth 		datetime ,
						@vSex 			varchar(1)  ,
						@vZipcode 		varchar(7) ,
						@vTelephone			varchar(14) = NULL ,
						@vHcode 		varchar(6) ,
						@iKatei			int ,		-- 課程は無視
						@iGakka		int ,		-- 学科も無視
						@vNengo		varchar(1) ,
						@iGradNendo  		int , 
						@iYobiko	int , -- genro = JukenNendo - SotugyoNendo
						@iZaigaku	int , -- 現役の大学生 genro->7
						@iSotugyo	int , -- 大学卒業　　　genro->5
						@iChutai	int , -- 大学中退     gento ->7
						@iBackgroundId	int , -- 労働者		genro->8
						@iLanguageSubject	int , -- 選択外国語
						@iScienceSub1	int , -- 選択理科
						@iScienceSub2	int , -- 選択理科
						@iPreferenceDay1Flag	int , -- 面接希望日１
						@iPreferenceDay2Flag	int , -- 面接希望日2
						@iPreferenceDay3Flag	int , -- 面接希望日3
						@iMenDate4	int , -- 面接希望日4
						@vMultipleApplyFlag	varchar(1) ,
						@vPrefectureName		varchar(12) = NULL ,
						@vCityName		varchar(18)=NULL ,
						@vAddress1	varchar(32)=NULL ,
						@vAddress2	varchar(32)=NULL ,
						@vAppato	varchar(32)=NULL
AS
declare	@iExamineeProfileID int,
	@dtBirthDay datetime ,
	@iSex int ,
	@iUniversityType int ,   -- Space
	@iQualificationId int ,    -- Space
	@iAdmissionType1 int,
	@iAdmissionType2 int ,
	@iZipCodeId int,
	@iHighSchoolId int ,
	@vAddress varchar(255),
	@vZipAddress varchar(255),
	@vPatAddress varchar(255) ,
	@iMultipleApplyFlag int ,
	@iLanguageSubjectProfileId int ,
	@iScienceSubProfileId1 int ,
	@iScienceSubProfileId2 int 
SELECT @vAddress=RTRIM(REPLACE(@vCityName,'　',' '))+RTRIM(REPLACE(@vAddress1,'　',' '))+RTRIM(REPLACE(@vAddress2,'　',' '))+RTRIM(REPLACE(@vAppato,'　',' '))
-- ProfileIDの最大値を求める
 SELECT  @iExamineeProfileID=MAX( iExamineeProfileID)+1 FROM tbSTEExamineeProfile
-- 和暦を変換する
  IF ( Datepart( year,@dtBirth) > 2045 )
    BEGIN
	SET @dtBirthDay=DATEADD( year ,-75, @dtBirth )
    END
  ELSE  IF ( Datepart( year,@dtBirth) > 2001 )
    BEGIN
	SET @dtBirthDay=DATEADD( year ,88, @dtBirth  )
    END
  ELSE
    BEGIN
	SET @dtBirthDay=DATEADD( year ,25, @dtBirth  )
    END
  IF ( @vMultipleApplyFlag = ' ' )
   SET @iMultipleApplyFlag=0
 ELSE
   SET @iMultipleApplyFlag=1
  IF ( @vSex = 'M' )
   SET @iSex=0
 ELSE
   SET @iSex=1
  SET @iUniversityType=100   -- Space
  SET @iQualificationId=500    -- Space
  IF ( @iZaigaku = 1 )
   BEGIN
	SET @iAdmissionType1=7
	SET @iUniversityType=101
   END
  ELSE IF ( @iSotugyo=1 )
   BEGIN
	SET @iAdmissionType1=5
	SET @iUniversityType=102
   END
  ELSE IF ( @iChutai=1 )
   BEGIN
	SET @iAdmissionType1=5
	SET @iUniversityType=103
   END
  ELSE
   BEGIN
  	IF ( @vNengo = 'H' )
		SELECT @iGradNendo=@iGradNendo+1988
 	ELSE 
		SELECT @iGradNendo=@iGradNendo+1925
	SELECT @iAdmissionType1=@iNendo-@iGradNendo
	SET @iAdmissionType2=@iAdmissionType1
   END
    -- 1:物理（１０）　２：化学（１１）　３：生物（１２）
  IF ( @iScienceSub1 = 1 )
	SET @iScienceSubProfileId1=14
  ELSE IF ( @iScienceSub1 = 2 )
	SET @iScienceSubProfileId1=15
  ELSE IF ( @iScienceSub1 = 3 )
	SET  @iScienceSubProfileId1=16
  IF ( @iScienceSub2 = 1 )
	SELECT @iScienceSubProfileId2=14
  ELSE IF ( @iScienceSub2 = 2 )
	SELECT @iScienceSubProfileId2=15
  ELSE IF ( @iScienceSub2 = 3 )
	SELECT @iScienceSubProfileId2=16
  IF ( @iLanguageSubject = 0 )
	SET @iLanguageSubjectProfileId=10
  ELSE IF ( @iLanguageSubject=1 )
	SET @iLanguageSubjectProfileId=11
  ELSE IF ( @iLanguageSubject=2 )
	SET @iLanguageSubjectProfileId=12
  ELSE
	SET @iLanguageSubjectProfileId=10
-- 高校コードを検索する。
   SET @iHighSchoolId=NULL
    SELECT @iHighSchoolId=iHighSchoolId FROM tbSTEHighSchoolType
	WHERE vHighSchoolCode=@vHcode
    IF( @iHighSchoolId IS  NULL )
    BEGIN
      SELECT @iHighSchoolId=MAX(iHighSchoolId)+1 FROM tbSTEHighSchoolType
      INSERT INTO  tbSTEHighSchoolType ( iHighSchoolId , vHighSchoolCode , dtCreate , dtUpdate)
		VALUES ( @iHighSchoolId , @vHcode , getdate() , getdate() )
    END
-- 郵便番号辞書を検索する。
--
-- 郵便番号を７桁にする。
--
SELECT @vZipCode=RTRIM(@vZipCode)
  WHILE LEN( @vZipCode ) < 7 
    BEGIN
      SELECT @vZipCode=@vZipCode+'0'
    END
    SET @iZipCodeId=NULL
    SELECT TOP 1 @iZipCodeId=iZipCodeId,@vZipaddress=vAddress1, @vPatAddress='%'+vAddress1+'%' FROM tbSTEZipCodeMaster
	WHERE vZipCodeName=@vZipCode ORDER BY vZipCodeName
SELECT @iZipCodeId,@vZipaddress, @vPatAddress,@vZipCode
    IF( @iZipCodeId IS NOT NULL )
    BEGIN
   --
   -- 自動インサートはしない。
   -- アドレスを最適な長さに切り詰める。（郵便番号辞書以降の文字列に切り詰める）
	IF ( PATINDEX(@vPatAddress,@vAddress) > 0 )
	      SELECT @vAddress=SUBSTRING(@vAddress,(LEN(@vZipaddress)+PATINDEX(@vPatAddress,@vAddress)) , 255)
    END
-- Insert new records into ExamineeTable
SET @iBackgroundId=200
insert into tbSTEExamineeProfile
 (
	iExamineeProfileID,
	iJukenNumber ,
	iNendo ,
	vExamineeName ,
	vKanaName ,
	vAddress ,
	iZipCodeId ,
	iSex ,
	iHighSchoolId ,
	vTelephone ,
	dtBirthDay ,
	iUniversityType ,
	iBackgroundId ,
	iQualificationId ,
	iLanguageSubjProfileId ,
	iScienceSubjProfileId1 ,
	iScienceSubjProfileId2 , 
	iPreferenceDay1Flag ,
	iPreferenceDay2Flag ,
	iPreferenceDay3Flag ,
	iMultipleApplyFlag ,
	iAdmissionType1 ,
	iAdmissionType2 ,
	dtCreate ,
	dtUpdate ,
	iExamineeStatus,
	iAbsentFlag ,
	iRejectFlag
 )
 VALUES 
(
	ISNULL(@iExamineeProfileID,1),
	@iJukenNumber ,
	@iNendo ,
	@vExamineeName ,
	@vKanaName ,
	@vAddress ,
	@iZipCodeId ,
	@iSex ,
	@iHighSchoolId ,
	@vTelephone ,
	@dtBirthDay ,
	@iUniversityType ,   -- Space
	@iBackgroundId ,
	@iQualificationId ,    -- Space
	@iLanguageSubjectProfileId ,
	@iScienceSubProfileId1 ,
	@iScienceSubProfileId2 , 
	@iPreferenceDay1Flag ,
	@iPreferenceDay2Flag ,
	@iPreferenceDay3Flag ,
	@iMultipleApplyFlag ,
	@iAdmissionType1 ,
	@iAdmissionType2 ,
	getdate() ,
	getdate() ,
	0 ,
	0 ,
	0
)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE    PROCEDURE uspSTESeisekiIchiran	 @iReportNo int ,
						 @iNendo int
AS
SET NOCOUNT ON
DECLARE	@iSerialNo int ,
		@iSpecialProfileId int ,
		@iSubjectProfileId int ,
		@iSeisekiIchiranId int ,
		@vSubjectName varchar(255) ,
		@vSubjectName01 varchar(255) ,
		@vSubjectName02 varchar(255) ,
		@vSubjectName03 varchar(255) ,
		@vSubjectName04 varchar(255) ,
		@vSubjectName05 varchar(255) ,
		@vSubjectName06 varchar(255) ,
		@vSubjectName07 varchar(255) ,
		@vSubjectName08 varchar(255) ,
		@vSubjectName09 varchar(255) ,
		@vSubjectName10 varchar(255) ,
		@vSubjectName11 varchar(255) ,
		@vSubjectName12 varchar(255) ,
		@vSubjectName13 varchar(255) ,
		@vSubjectName14 varchar(255) ,
		@vSubjectName15 varchar(255) ,
		@vSubjectName16 varchar(255) ,
		@vSubjectName17 varchar(255) ,
		@vSubjectName18 varchar(255) ,
		@vSubjectName19 varchar(255) ,
		@vSubjectName20 varchar(255) ,
		@vSubjectName21 varchar(255) ,
		@vSubjectName22 varchar(255) ,
		@vSubjectName23 varchar(255) ,
		@vSubjectName24 varchar(255) ,
		@vSubjectName25 varchar(255) ,
		@vSubjectName26 varchar(255) ,
		@vSubjectName27 varchar(255) ,
		@vSubjectName28 varchar(255) ,
		@vSubjectName29 varchar(255) ,
		@vSubjectName30 varchar(255) ,
		@vSubjectName31 varchar(255) ,
		@vSubjectName32 varchar(255) ,
		@vSubjectName33 varchar(255) ,
		@vSubjectName34 varchar(255) ,
		@vSubjectName35 varchar(255) 
CREATE TABLE #wk_tbl01 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float , vSubject varchar(64) , vMark varchar(40) , vMarkBSL varchar(40) )
CREATE TABLE #wk_tbl02 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl03 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl04 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl05 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl06 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl07 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl08 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl09 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl10 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl11 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl12 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl13 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl14 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl15 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl16 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl17 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl18 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl19 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl20 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl21 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl22 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl23 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl24 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl25 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl26 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl27 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl28 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl29 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl30 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl31 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl32 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl33 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl34 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
CREATE TABLE #wk_tbl35 (  iID int IDENTITY (1, 1) NOT NULL , iExamineeProfileId int , fRawScore float, fBSLScore float  , vSubject varchar(64)  , vMark varchar(40) , vMarkBSL varchar(40))
DECLARE crSeisekiPrint CURSOR
FOR
  SELECT iReportNo,iSerialNo,iSpecialProfileID,iSubjectProfileId,iSeisekiIchiranId
	FROM tbSTESeisekiIchiranProfile
	WHERE iReportNo=@iReportNo
OPEN crSeisekiPrint
  
 FETCH  NEXT FROM crSeisekiPrint INTO @iReportNo,@iSerialNo,@iSpecialProfileID,@iSubjectProfileId,@iSeisekiIchiranId 
 IF ( @@FETCH_STATUS = 0 )
   BEGIN
      WHILE @@FETCH_STATUS = 0
         BEGIN
	SELECT @vSubjectName=vButtonName FROM tbSTESeisekiSpecialProfile
		WHERE iSpecialProfileId=@iSpecialProfileID
	 IF 		( @iSerialNo=1 ) INSERT INTO #wk_tbl01 ( iExamineeProfileId , fRawScore , fBSLScore , vSubject ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName01 OUTPUT
           ELSE IF 	( @iSerialNo=2 ) INSERT INTO #wk_tbl02 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName02 OUTPUT
           ELSE IF 	( @iSerialNo=3 ) INSERT INTO #wk_tbl03 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName03 OUTPUT
           ELSE IF 	( @iSerialNo=4 ) INSERT INTO #wk_tbl04 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName04 OUTPUT
           ELSE IF 	( @iSerialNo=5 ) INSERT INTO #wk_tbl05 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName05 OUTPUT
           ELSE IF 	( @iSerialNo=6 ) INSERT INTO #wk_tbl06 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName06 OUTPUT
           ELSE IF 	( @iSerialNo=7 ) INSERT INTO #wk_tbl07 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName07 OUTPUT
           ELSE IF 	( @iSerialNo=8 ) INSERT INTO #wk_tbl08 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName08 OUTPUT
           ELSE IF 	( @iSerialNo=9 ) INSERT INTO #wk_tbl09 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName09 OUTPUT
           ELSE IF 	( @iSerialNo=10 ) INSERT INTO #wk_tbl10 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName10 OUTPUT
           ELSE IF 	( @iSerialNo=11 ) INSERT INTO #wk_tbl11 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName11 OUTPUT
           ELSE IF 	( @iSerialNo=12 ) INSERT INTO #wk_tbl12 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName12 OUTPUT
           ELSE IF 	( @iSerialNo=13 ) INSERT INTO #wk_tbl13 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName13 OUTPUT
           ELSE IF 	( @iSerialNo=14 ) INSERT INTO #wk_tbl14 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName14 OUTPUT
           ELSE IF 	( @iSerialNo=15 ) INSERT INTO #wk_tbl15 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName15 OUTPUT
           ELSE IF 	( @iSerialNo=16 ) INSERT INTO #wk_tbl16 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName16 OUTPUT
           ELSE IF 	( @iSerialNo=17 ) INSERT INTO #wk_tbl17 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName17 OUTPUT
           ELSE IF 	( @iSerialNo=18 ) INSERT INTO #wk_tbl18 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName18 OUTPUT
           ELSE IF 	( @iSerialNo=19 ) INSERT INTO #wk_tbl19 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName19 OUTPUT
           ELSE IF 	( @iSerialNo=20 ) INSERT INTO #wk_tbl20 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName20 OUTPUT
           ELSE IF 	( @iSerialNo=21 ) INSERT INTO #wk_tbl21 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName21 OUTPUT
           ELSE IF 	( @iSerialNo=22 ) INSERT INTO #wk_tbl22 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName22 OUTPUT
           ELSE IF 	( @iSerialNo=23 ) INSERT INTO #wk_tbl23 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName23 OUTPUT
           ELSE IF 	( @iSerialNo=24 ) INSERT INTO #wk_tbl24 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName24 OUTPUT
           ELSE IF 	( @iSerialNo=25 ) INSERT INTO #wk_tbl25 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName25 OUTPUT
           ELSE IF 	( @iSerialNo=26 ) INSERT INTO #wk_tbl26 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName26 OUTPUT
           ELSE IF 	( @iSerialNo=27 ) INSERT INTO #wk_tbl27 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName27 OUTPUT
           ELSE IF 	( @iSerialNo=28 ) INSERT INTO #wk_tbl28 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName28 OUTPUT
           ELSE IF 	( @iSerialNo=29 ) INSERT INTO #wk_tbl29 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName29 OUTPUT
           ELSE IF 	( @iSerialNo=30 ) INSERT INTO #wk_tbl30 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName30 OUTPUT
           ELSE IF 	( @iSerialNo=31 ) INSERT INTO #wk_tbl31 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName31 OUTPUT
           ELSE IF 	( @iSerialNo=32 ) INSERT INTO #wk_tbl32 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName32 OUTPUT
           ELSE IF 	( @iSerialNo=33 ) INSERT INTO #wk_tbl33 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName33 OUTPUT
           ELSE IF 	( @iSerialNo=34 ) INSERT INTO #wk_tbl34 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName34 OUTPUT
           ELSE IF 	( @iSerialNo=35 ) INSERT INTO #wk_tbl35 ( iExamineeProfileId , fRawScore ,fBSLScore , vSubject  ,vMark , vMarkBSL ) EXEC uspSTESeisekiStudentScore @iSpecialProfileID,@iSubjectProfileId,@vSubjectName,@iSeisekiIchiranId , @vSubjectName35 OUTPUT
	 FETCH  NEXT FROM crSeisekiPrint INTO @iReportNo,@iSerialNo,@iSpecialProfileID,@iSubjectProfileId,@iSeisekiIchiranId 
         END
   END
 
CLOSE crSeisekiPrint
DEALLOCATE crSeisekiPrint
/*
 *
 */
DELETE tbSTETempForPrint WHERE @iReportNo=iReportNo AND @iNendo=iReportSubNo
INSERT INTO tbSTETempForPrint
SELECT @iReportNo iReportNo ,
	 @iNendo iReportSubNo ,
	 vwSTEExaminee.iExamineeProfileId ,
	 vwSTEExaminee.iJukenNumber ,
	 vwSTEExaminee.iNendo ,
	 vwSTEExaminee.vExamineeName ,
	 vwSTEExaminee.vKanaName ,
	 vwSTEExaminee.iSex ,
	 vwSTEExaminee.dtBirthDay ,
	 vwSTEExaminee.iAbsentFlag ,
	 vwSTEExaminee.iRejectFlag ,
	 vwSTEExaminee.iExamineeStatus ,
	 vwSTEExaminee.dtSecondExamDay ,
	 vwSTEExaminee.iUniversityType ,
	 vwSTEExaminee.iBackgroundId ,
	 vwSTEExaminee.iFamilyId ,
	 vwSTEExaminee.iParentJobCategory ,
	 vwSTEExaminee.iQualificationId ,
	 vwSTEExaminee.iSuisenFlagId ,
	 vwSTEExaminee.vNationality ,
	 vwSTEExaminee.iPhysicalConditionId ,
	 vwSTEExaminee.iLanguageSubjProfileId ,
	 vwSTEExaminee.iScienceSubjProfileId1 ,
	 vwSTEExaminee.iScienceSubjProfileId2 ,
	 vwSTEExaminee.iPreferenceDay1Flag ,
	 vwSTEExaminee.iPreferenceDay2Flag ,
	 vwSTEExaminee.iPreferenceDay3Flag ,
	 vwSTEExaminee.iMultipleApplyFlag ,
	 vwSTEExaminee.iAdmissionType1 ,
	 vwSTEExaminee.iRandom ,
	 vwSTEExaminee.vHighSchoolCode ,
	 vwSTEExaminee.iAge ,
	(CASE  vwSTEExaminee.vNationality 
		WHEN '日本'  THEN 0 
		WHEN  NULL THEN 0 
		WHEN '' THEN 0 
		ELSE 1 
	END) iNationality ,
	  #wk_tbl01.fRawScore fRawScore01 , 
	  #wk_tbl01.fBSLScore fChoseiScore01 , 
	  @vSubjectName01    vSubject01 , 
	  #wk_tbl01.vMark       vUpMark01 , 
	  #wk_tbl01.vMarkBSL  vLowMark01 , 
	  #wk_tbl02.fRawScore fRawScore02 , 
	  #wk_tbl02.fBSLScore fChoseiScore02 , 
	  @vSubjectName02    vSubject02 , 
	  #wk_tbl02.vMark       vUpMark02 , 
	  #wk_tbl02.vMarkBSL  vLowMark02 , 
	  #wk_tbl03.fRawScore fRawScore03 , 
	  #wk_tbl03.fBSLScore fChoseiScore03 , 
	 @vSubjectName03    vSubject03 , 
	  #wk_tbl03.vMark       vUpMark03 , 
	  #wk_tbl03.vMarkBSL  vLowMark03 , 
	  #wk_tbl04.fRawScore fRawScore04 , 
	  #wk_tbl04.fBSLScore fChoseiScore04 , 
	  @vSubjectName04    vSubject04 , 
	  #wk_tbl04.vMark       vUpMark04 , 
	  #wk_tbl04.vMarkBSL  vLowMark04 , 
	  #wk_tbl05.fRawScore fRawScore05 , 
	  #wk_tbl05.fBSLScore fChoseiScore05 , 
	  @vSubjectName05    vSubject05 , 
	  #wk_tbl05.vMark       vUpMark05 , 
	  #wk_tbl05.vMarkBSL  vLowMark05 , 
	  #wk_tbl06.fRawScore fRawScore06 , 
	  #wk_tbl06.fBSLScore fChoseiScore06 , 
	  @vSubjectName06    vSubject06 , 
	  #wk_tbl06.vMark       vUpMark06 , 
	  #wk_tbl06.vMarkBSL  vLowMark06 , 
	  #wk_tbl07.fRawScore fRawScore07 , 
	  #wk_tbl07.fBSLScore fChoseiScore07 , 
	  @vSubjectName07    vSubject07 , 
	  #wk_tbl07.vMark       vUpMark07 , 
	  #wk_tbl07.vMarkBSL  vLowMark07 , 
	  #wk_tbl08.fRawScore fRawScore08 , 
	  #wk_tbl08.fBSLScore fChoseiScore08 , 
	  @vSubjectName08    vSubject08 , 
	  #wk_tbl08.vMark       vUpMark08 , 
	  #wk_tbl08.vMarkBSL  vLowMark08 , 
	  #wk_tbl09.fRawScore fRawScore09 , 
	  #wk_tbl09.fBSLScore fChoseiScore09 , 
	  @vSubjectName09    vSubject09 , 
	  #wk_tbl09.vMark       vUpMark09 , 
	  #wk_tbl09.vMarkBSL  vLowMark09 , 
	  #wk_tbl10.fRawScore fRawScore10 , 
	  #wk_tbl10.fBSLScore fChoseiScore10 , 
	  @vSubjectName10    vSubject10 , 
	  #wk_tbl10.vMark       vUpMark10 , 
	  #wk_tbl10.vMarkBSL  vLowMark10 , 
	  #wk_tbl11.fRawScore fRawScore11 , 
	  #wk_tbl11.fBSLScore fChoseiScore11 , 
	  @vSubjectName11    vSubject11 , 
	  #wk_tbl11.vMark       vUpMark11 , 
	  #wk_tbl11.vMarkBSL  vLowMark11 , 
	  #wk_tbl12.fRawScore fRawScore12 , 
	  #wk_tbl12.fBSLScore fChoseiScore12 , 
	  @vSubjectName12    vSubject12 , 
	  #wk_tbl12.vMark       vUpMark12 , 
	  #wk_tbl12.vMarkBSL  vLowMark12 , 
	  #wk_tbl13.fRawScore fRawScore13 , 
	  #wk_tbl13.fBSLScore fChoseiScore13 , 
	  @vSubjectName13    vSubject13 , 
	  #wk_tbl13.vMark       vUpMark13 , 
	  #wk_tbl13.vMarkBSL  vLowMark13 , 
	  #wk_tbl14.fRawScore fRawScore14 , 
	  #wk_tbl14.fBSLScore fChoseiScore14 , 
	  @vSubjectName14    vSubject14 , 
	  #wk_tbl14.vMark       vUpMark14 , 
	  #wk_tbl14.vMarkBSL  vLowMark14 , 
	  #wk_tbl15.fRawScore fRawScore15 , 
	  #wk_tbl15.fBSLScore fChoseiScore15 , 
	  @vSubjectName15    vSubject15 , 
	  #wk_tbl15.vMark       vUpMark15 , 
	  #wk_tbl15.vMarkBSL  vLowMark15 , 
	  #wk_tbl16.fRawScore fRawScore16 , 
	  #wk_tbl16.fBSLScore fChoseiScore16 , 
	  @vSubjectName16    vSubject16 , 
	  #wk_tbl16.vMark       vUpMark16 , 
	  #wk_tbl16.vMarkBSL  vLowMark16 , 
	  #wk_tbl17.fRawScore fRawScore17 , 
	  #wk_tbl17.fBSLScore fChoseiScore17 , 
	  @vSubjectName17    vSubject17 , 
	  #wk_tbl17.vMark       vUpMark17 , 
	  #wk_tbl17.vMarkBSL  vLowMark17 , 
	  #wk_tbl18.fRawScore fRawScore18 , 
	  #wk_tbl18.fBSLScore fChoseiScore18 , 
	  @vSubjectName18    vSubject18 , 
	  #wk_tbl18.vMark       vUpMark18 , 
	  #wk_tbl18.vMarkBSL  vLowMark18 , 
	  #wk_tbl19.fRawScore fRawScore19 , 
	  #wk_tbl19.fBSLScore fChoseiScore19 , 
	  @vSubjectName19    vSubject19 , 
	  #wk_tbl19.vMark       vUpMark19 , 
	  #wk_tbl19.vMarkBSL  vLowMark19 , 
	  #wk_tbl20.fRawScore fRawScore20 , 
	  #wk_tbl20.fBSLScore fChoseiScore20 , 
	  @vSubjectName21    vSubject20 , 
	  #wk_tbl20.vMark       vUpMark20 , 
	  #wk_tbl20.vMarkBSL  vLowMark20 , 
	  #wk_tbl21.fRawScore fRawScore21 , 
	  #wk_tbl21.fBSLScore fChoseiScore21 , 
	  #wk_tbl21.vSubject    vSubject21 , 
	  #wk_tbl21.vMark       vUpMark21 , 
	  #wk_tbl21.vMarkBSL  vLowMark21 , 
	  #wk_tbl22.fRawScore fRawScore22 , 
	  #wk_tbl22.fBSLScore fChoseiScore22 , 
	  @vSubjectName22    vSubject22 , 
	  #wk_tbl22.vMark       vUpMark22 , 
	  #wk_tbl22.vMarkBSL  vLowMark22 , 
	  #wk_tbl23.fRawScore fRawScore23 , 
	  #wk_tbl23.fBSLScore fChoseiScore23 , 
	  @vSubjectName23    vSubject23 , 
	  #wk_tbl23.vMark       vUpMark23 , 
	  #wk_tbl23.vMarkBSL  vLowMark23 , 
	  #wk_tbl24.fRawScore fRawScore24 , 
	  #wk_tbl24.fBSLScore fChoseiScore24 , 
	  @vSubjectName24    vSubject24 , 
	  #wk_tbl24.vMark       vUpMark24 , 
	  #wk_tbl24.vMarkBSL  vLowMark24 , 
	  #wk_tbl25.fRawScore fRawScore25 , 
	  #wk_tbl25.fBSLScore fChoseiScore25 , 
	  @vSubjectName25    vSubject25 , 
	  #wk_tbl25.vMark       vUpMark25 , 
	  #wk_tbl25.vMarkBSL  vLowMark25 , 
	  #wk_tbl26.fRawScore fRawScore26 , 
	  #wk_tbl26.fBSLScore fChoseiScore26 , 
	  @vSubjectName26    vSubject26 , 
	  #wk_tbl26.vMark       vUpMark26 , 
	  #wk_tbl26.vMarkBSL  vLowMark26 , 
	  #wk_tbl27.fRawScore fRawScore27 , 
	  #wk_tbl27.fBSLScore fChoseiScore27 , 
	 @vSubjectName27    vSubject27 , 
	  #wk_tbl27.vMark       vUpMark27 , 
	  #wk_tbl27.vMarkBSL  vLowMark27 , 
	  #wk_tbl28.fRawScore fRawScore28 , 
	  #wk_tbl28.fBSLScore fChoseiScore28 , 
	 @vSubjectName28    vSubject28 , 
	  #wk_tbl28.vMark       vUpMark28 , 
	  #wk_tbl28.vMarkBSL  vLowMark28 , 
	  #wk_tbl29.fRawScore fRawScore29 , 
	  #wk_tbl29.fBSLScore fChoseiScore29 , 
	  @vSubjectName29    vSubject29 , 
	  #wk_tbl29.vMark       vUpMark29 , 
	  #wk_tbl29.vMarkBSL  vLowMark29 , 
	  #wk_tbl30.fRawScore fRawScore30 , 
	  #wk_tbl30.fBSLScore fChoseiScore30 , 
	  @vSubjectName30    vSubject30 , 
	  #wk_tbl30.vMark       vUpMark30 , 
	  #wk_tbl30.vMarkBSL  vLowMark30 , 
	  #wk_tbl31.fRawScore fRawScore31 , 
	  #wk_tbl31.fBSLScore fChoseiScore31 , 
	  @vSubjectName31    vSubject31 , 
	  #wk_tbl31.vMark       vUpMark31 , 
	  #wk_tbl31.vMarkBSL  vLowMark31 , 
	  #wk_tbl32.fRawScore fRawScore32 , 
	  #wk_tbl32.fBSLScore fChoseiScore32 , 
	  @vSubjectName32    vSubject32 , 
	  #wk_tbl32.vMark       vUpMark32 , 
	  #wk_tbl32.vMarkBSL  vLowMark32 , 
	  #wk_tbl33.fRawScore fRawScore33 , 
	  #wk_tbl33.fBSLScore fChoseiScore33 , 
	  @vSubjectName33    vSubject33 , 
	  #wk_tbl33.vMark       vUpMark33 , 
	  #wk_tbl33.vMarkBSL  vLowMark33 , 
	  #wk_tbl34.fRawScore fRawScore34 , 
	  #wk_tbl34.fBSLScore fChoseiScore34 , 
	  @vSubjectName34    vSubject34 , 
	  #wk_tbl34.vMark       vUpMark34 , 
	  #wk_tbl34.vMarkBSL  vLowMark34 , 
	  #wk_tbl35.fRawScore fRawScore35 , 
	  #wk_tbl35.fBSLScore fChoseiScore35 , 
	  @vSubjectName35    vSubject35 , 
	  #wk_tbl35.vMark       vUpMark35 , 
	  #wk_tbl35.vMarkBSL  vLowMark35 
	FROM
  		vwSTEExaminee 
		   LEFT OUTER JOIN
                      #wk_tbl01 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl01.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl02 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl02.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl03 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl03.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl04 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl04.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl05 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl05.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl06 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl06.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl07 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl07.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl08 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl08.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl09 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl09.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl10 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl10.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl11 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl11.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl12 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl12.iExamineeProfileId
                       LEFT OUTER JOIN
                     #wk_tbl13 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl13.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl14 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl14.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl15 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl15.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl16 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl16.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl17 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl17.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl18 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl18.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl19 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl19.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl20 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl20.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl21 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl21.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl22 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl22.iExamineeProfileId
                       LEFT OUTER JOIN
                     #wk_tbl23 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl23.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl24 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl24.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl25 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl25.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl26 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl26.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl27 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl27.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl28 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl28.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl29 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl29.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl30 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl30.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl31 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl31.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl32 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl32.iExamineeProfileId
                       LEFT OUTER JOIN
                     #wk_tbl33 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl33.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl34 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl34.iExamineeProfileId
                       LEFT OUTER JOIN
                      #wk_tbl35 ON 
                      vwSTEExaminee.iExamineeProfileId=#wk_tbl35.iExamineeProfileId
	WHERE  vwSTEExaminee.iNendo=@iNendo

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/*
 *
 *
 *
 */
CREATE      PROCEDURE uspSTESeisekiStudentScore	@iSpecialProfileID	int ,
							@iSubjectProfileId int ,
							@vSpecialSubjectName	varchar(255) , 
							@iSeisekiIchiranId int ,
							@vSubjectName varchar(64) OUTPUT
AS
-- DECLARE @vSubjectName varchar(64)
 SELECT @vSubjectName=vSubjectName FROM tbSTESubjectProfile
	WHERE iSubjectProfileId=@iSubjectProfileId
 SELECT @vSubjectName=ISNULL(  @vSubjectName , @vSpecialSubjectName )
-- SELECT @@vSubjectName=@vSubjectName
IF ( @iSpecialProfileID IS NULL )
  BEGIN
	-- 科目ごとの点数を上位よりソートして返す。
	SELECT	iExamineeProfileId ,
			fRawScore fRawScore ,
			iAbsentFlag fBSLScore ,
			@vSubjectName vSubject ,
			'' mark1 ,
			'' mark2
		FROM vwSTEScoreProfile
		WHERE (iActiveFlag=1 AND iSystemNendo=iNendo )AND iSubjectProfileId = @iSubjectProfileId 
--		GROUP BY iExamineeProfileId
		ORDER BY fRawScore DESC
  END
ELSE  IF( @iSpecialProfileID in ( 100,103 ) )
  BEGIN
-- 100,"一次素点",100,,
-- 103,"二次素点",103,,
	SELECT	iExamineeProfileId ,
			SUM(fRawScore) fRawScore ,
			SUM(fChoseiScore) fBSLScore ,
			@vSubjectName vSubject ,
			'' mark1 ,
			'' mark2
		FROM vwSTEScoreProfile
		WHERE (iActiveFlag=1 AND iSystemNendo=iNendo ) AND iSubjectProfileId in (SELECT iSubjectProfileId FROM tbSTESeisekiSpecialSubjectProfile WHERE iSeisekiIchiranId=@iSeisekiIchiranId)
		GROUP BY iExamineeProfileId
		ORDER BY SUM(fRawScore) DESC
  END
ELSE  IF( @iSpecialProfileID in ( 102,106,107 ) )
  BEGIN
-- 102,"一次合計",102,,
-- 106,"二次合計",106,,
-- 107,"総合計",107,,
	SELECT	iExamineeProfileId ,
			SUM(fRawScore+fChoseiScore) fRawScore ,
			SUM(fChoseiScore) fBSLScore ,
			@vSubjectName vSubject ,
			'' mark1 ,
			'' mark2
		FROM vwSTEScoreProfile
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo ) AND iSubjectProfileId in (SELECT iSubjectProfileId FROM tbSTESeisekiSpecialSubjectProfile WHERE iSeisekiIchiranId=@iSeisekiIchiranId)
		GROUP BY iExamineeProfileId
		ORDER BY SUM(fRawScore+fChoseiScore) DESC
  END
ELSE  IF( @iSpecialProfileID in ( 101,104,105,110,111,112,120,121,122,123,124,125,126,127,130,131 ) )
  BEGIN
-- 101,"一次調整",101,,
-- 104,"二次調整",104,,
-- 105,"調整合計",105,,
-- 110,"評定調整",110,,
-- 111,"欠席調整",111,,
-- 112,"現浪調整",112,,
-- 120,"英語調整",120,,
-- 121,"独語調整",121,,
-- 122,"仏語調整",122,,
-- 123,"数学調整",123,,
-- 124,"物理調整",124,,
-- 125,"化学調整",125,,
-- 126,"生物調整",126,,
-- 130,"小論調整",130,,
-- 131,"面接調整",131,,
	SELECT	iExamineeProfileId ,
			SUM(fChoseiScore) fRawScore ,
			SUM(fRawScore) fBSLScore ,
			@vSubjectName vSubject ,
			'' mark1 ,
			'' mark2
		FROM vwSTEScoreProfile
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo ) AND iSubjectProfileId in (SELECT iSubjectProfileId FROM tbSTESeisekiSpecialSubjectProfile WHERE iSeisekiIchiranId=@iSeisekiIchiranId)
		GROUP BY iExamineeProfileId
		ORDER BY SUM(fChoseiScore) DESC
  END
ELSE  IF( @iSpecialProfileID = 200  )
  BEGIN
-- 200,"性別年齢",200,,
	SELECT	iExamineeProfileId ,
			iSex  fRawScore ,
			iAge  fBSLScore ,
			@vSubjectName vSubject ,
			vSex  mark1 ,
			CONVERT(varchar,iAge) mark2
		FROM vwSTRExaminee
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo )
--		GROUP BY iExamineeProfileId
		ORDER BY iExamineeProfileId
  END
ELSE  IF( @iSpecialProfileID = 201  )
  BEGIN
/*
 201,"現浪区分",201,,
*/
	SELECT	iExamineeProfileId ,
			iAdmissionType1  fRawScore ,
			0  fBSLScore ,
			@vSubjectName vSubject ,
			CONVERT(varchar,iAdmissionType1)  mark1 ,
			'' mark2
		FROM vwSTRExaminee
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo )
--		GROUP BY iExamineeProfileId
		ORDER BY iExamineeProfileId
  END
ELSE  IF( @iSpecialProfileID = 202  )
  BEGIN
/*
  202,"成績概評",202,,
*/
	SELECT	iExamineeProfileId ,
			0  fRawScore ,
			0  fBSLScore ,
			@vSubjectName vSubject ,
			vHyoteiGrade  mark1 ,
			'' mark2
		FROM vwSTRExaminee
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo )
--		GROUP BY iExamineeProfileId
		ORDER BY iExamineeProfileId
  END
ELSE  IF( @iSpecialProfileID = 203  )
  BEGIN
/*
 203,"父母子弟",203,,
*/
	SELECT	iExamineeProfileId ,
			iFamilyId  fRawScore ,
			iParentJobCategory  fBSLScore ,
			@vSubjectName vSubject ,
			vFamily  mark1 ,
			vParentJob mark2
		FROM vwSTRExaminee
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo )
--		GROUP BY iExamineeProfileId
		ORDER BY iExamineeProfileId
  END
ELSE  IF( @iSpecialProfileID = 204  )
  BEGIN
/*
 204,"選択外語",204,,
*/
	SELECT	iExamineeProfileId ,
			0  fRawScore ,
			0  fBSLScore ,
			@vSubjectName vSubject ,
			vLangSubjectName  mark1 ,
			'' mark2
		FROM vwSTRExaminee
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo )
--		GROUP BY iExamineeProfileId
		ORDER BY iExamineeProfileId
  END
ELSE  IF( @iSpecialProfileID = 205  )
  BEGIN
/*
  205,"選択理科",205,,
*/
/*
 204,"選択外語",204,,
*/
	SELECT	iExamineeProfileId ,
			0  fRawScore ,
			0  fBSLScore ,
			@vSubjectName vSubject ,
			vRika1SubjectName  mark1 ,
			vRika2SubjectName mark2
		FROM vwSTRExaminee
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo )
--		GROUP BY iExamineeProfileId
		ORDER BY iExamineeProfileId
  END
ELSE  IF( @iSpecialProfileID = 206  )
  BEGIN
/*
  206,"二次試験",206,,
*/
	SELECT	iExamineeProfileId ,
			DATEPART(dd,dtSecondExamDay)  fRawScore ,
			0  fBSLScore ,
			@vSubjectName vSubject ,
			CONVERT(varchar,DATEPART(dd,dtSecondExamDay))  mark1 ,
			'' mark2
		FROM vwSTRExaminee
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo )
--		GROUP BY iExamineeProfileId
		ORDER BY iExamineeProfileId
  END
ELSE  IF( @iSpecialProfileID = 207  )
  BEGIN
/*
  207,"高校情報",207,,
*/
	SELECT	iExamineeProfileId ,
			0  fRawScore ,
			0  fBSLScore ,
			@vSubjectName vSubject ,
			(REPLACE(vHighSchoolName ,'　' ,'' )+'('+vHighSchoolCode+')')  mark1 ,
			vHPrefectureName mark2
		FROM vwSTRExaminee
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo )
--		GROUP BY iExamineeProfileId
		ORDER BY iExamineeProfileId
  END
ELSE  IF( @iSpecialProfileID = 208  )
  BEGIN
/*
208,"外国籍",208,,
*/
	SELECT	iExamineeProfileId ,
			(CASE  vNationality 
				WHEN '日本'  THEN 0 
				WHEN  NULL THEN 0 
				WHEN '' THEN 0 
				ELSE 1 
			END) fRawScore ,
			0  fBSLScore ,
			@vSubjectName vSubject ,
			vNationality  mark1 ,
			'' mark2
		FROM vwSTRExaminee
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo )
--		GROUP BY iExamineeProfileId
		ORDER BY iExamineeProfileId
  END
ELSE  IF( @iSpecialProfileID = 209  )
  BEGIN
/*
 209,"推薦",209,,
*/
	SELECT	iExamineeProfileId ,
			iSuisenFlagId  fRawScore ,
			0  fBSLScore ,
			@vSubjectName vSubject ,
			(vSuisen)  mark1 ,
			'' mark2
		FROM vwSTRExaminee
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo )
--		GROUP BY iExamineeProfileId
		ORDER BY iExamineeProfileId
  END
ELSE  IF( @iSpecialProfileID = 210  )
  BEGIN
/*
 210,"職歴",210,,
*/
	SELECT	iExamineeProfileId ,
			iBackgroundId  fRawScore ,
			0  fBSLScore ,
			@vSubjectName vSubject ,
			vBackground  mark1 ,
			'' mark2
		FROM vwSTRExaminee
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo )
--		GROUP BY iExamineeProfileId
		ORDER BY iExamineeProfileId
  END
ELSE  IF( @iSpecialProfileID = 211  )
  BEGIN
/*
 211,"大卒中退",211,,
*/
	SELECT	iExamineeProfileId ,
			iUniversityType  fRawScore ,
			0  fBSLScore ,
			@vSubjectName vSubject ,
			vUnivType  mark1 ,
			vUnivName mark2
		FROM vwSTRExaminee
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo )
--		GROUP BY iExamineeProfileId
		ORDER BY iExamineeProfileId
  END
ELSE  IF( @iSpecialProfileID = 212  )
  BEGIN
/*
 212,"状態更新１",212,,
*/
	SELECT	vwSTRExaminee.iExamineeProfileId ,
			tbSTEExamineeStatusTrail.iExamineeStatus  fRawScore ,
			tbSTEExamineeStatusTrail.iRejectFlag  fBSLScore ,
			@vSubjectName vSubject ,
			CONVERT(varchar,tbSTEExamineeStatusTrail.dtModify,111)  mark1 ,
			'' mark2
		FROM vwSTRExaminee LEFT OUTER JOIN tbSTEExamineeStatusTrail
			ON vwSTRExaminee.iExamineeProfileId=tbSTEExamineeStatusTrail.iExamineeProfileId
			AND tbSTEExamineeStatusTrail.iPos=0
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo ) 
--		GROUP BY iExamineeProfileId
		ORDER BY vwSTRExaminee.iExamineeProfileId
  END
ELSE  IF( @iSpecialProfileID = 213  )
  BEGIN
/*
  213,"状態更新２",213,,
*/
	SELECT	vwSTRExaminee.iExamineeProfileId ,
			tbSTEExamineeStatusTrail.iExamineeStatus  fRawScore ,
			tbSTEExamineeStatusTrail.iRejectFlag  fBSLScore ,
			@vSubjectName vSubject ,
			CONVERT(varchar,tbSTEExamineeStatusTrail.dtModify,111)  mark1 ,
			'' mark2
		FROM vwSTRExaminee LEFT OUTER JOIN tbSTEExamineeStatusTrail
			ON vwSTRExaminee.iExamineeProfileId=tbSTEExamineeStatusTrail.iExamineeProfileId
			AND tbSTEExamineeStatusTrail.iPos=1
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo ) 
--		GROUP BY iExamineeProfileId
		ORDER BY vwSTRExaminee.iExamineeProfileId
  END
ELSE  IF( @iSpecialProfileID = 214  )
  BEGIN
/*
  214,"状態更新３",214,,
*/
	SELECT	vwSTRExaminee.iExamineeProfileId ,
			tbSTEExamineeStatusTrail.iExamineeStatus  fRawScore ,
			tbSTEExamineeStatusTrail.iRejectFlag  fBSLScore ,
			@vSubjectName vSubject ,
			CONVERT(varchar,tbSTEExamineeStatusTrail.dtModify,111)  mark1 ,
			'' mark2
		FROM vwSTRExaminee LEFT OUTER JOIN tbSTEExamineeStatusTrail
			ON vwSTRExaminee.iExamineeProfileId=tbSTEExamineeStatusTrail.iExamineeProfileId
			AND tbSTEExamineeStatusTrail.iPos=2
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo ) 
--		GROUP BY iExamineeProfileId
		ORDER BY vwSTRExaminee.iExamineeProfileId
  END
ELSE  IF( @iSpecialProfileID = 215  )
  BEGIN
/*
  215,"状態更新４",215,,
*/
	SELECT	vwSTRExaminee.iExamineeProfileId ,
			tbSTEExamineeStatusTrail.iExamineeStatus  fRawScore ,
			tbSTEExamineeStatusTrail.iRejectFlag  fBSLScore ,
			@vSubjectName vSubject ,
			CONVERT(varchar,tbSTEExamineeStatusTrail.dtModify,111)  mark1 ,
			'' mark2
		FROM vwSTRExaminee LEFT OUTER JOIN tbSTEExamineeStatusTrail
			ON vwSTRExaminee.iExamineeProfileId=tbSTEExamineeStatusTrail.iExamineeProfileId
			AND tbSTEExamineeStatusTrail.iPos=3
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo ) 
  END
ELSE  IF( @iSpecialProfileID = 216  )
  BEGIN
/*
  216,"一次乱数",216,,
*/
	SELECT	iExamineeProfileId ,
			iRandom  fRawScore ,
			0  fBSLScore ,
			@vSubjectName vSubject ,
			CONVERT(varchar,iRandom)  mark1 ,
			'' mark2
		FROM vwSTRExaminee
		WHERE  (iActiveFlag=1 AND iSystemNendo=iNendo )
--		GROUP BY iExamineeProfileId
		ORDER BY iExamineeProfileId
  END
ELSE 
  BEGIN
	SELECT	iExamineeProfileId ,
			SUM(fRawScore) fRawScore ,
			SUM(fChoseiScore) fBSLScore ,
			@vSubjectName vSubject ,
			'' mark1 ,
			'' mark2
		FROM tbSTEScoreProfile
		WHERE iSubjectProfileId in (SELECT iSubjectProfileId FROM tbSTRSeisekiSpecialSubjectProfile WHERE iSeisekiIchiranId=@iSeisekiIchiranId)
		GROUP BY iExamineeProfileId
		ORDER BY SUM(fRawScore) DESC
  END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

/*
 *
 *
 *
 */
CREATE PROCEDURE uspSTRSeisekiStudentScore	@iSpecialProfileID	int ,
							@iSubjectGradeProfileId int ,
							@vSpecialSubjectName	varchar(255) , 
							@iSeisekiIchiranId int 
AS
 DECLARE @vSubjectName varchar(64)
 SELECT @vSubjectName=vSubjectName FROM vwSTRStudentSubjectScore
	WHERE iSubjectGradeProfileId=@iSubjectGradeProfileId
 SELECT @vSubjectName=ISNULL(  @vSubjectName , @vSpecialSubjectName )
IF ( @iSpecialProfileID IS NULL )
  BEGIN
	-- 科目ごとの点数を上位よりソートして返す。
	SELECT	iStudentProfileId ,
			SUM(fRawScore) fRawScore ,
			SUM(fGraceScore) fBSLScore ,
			@vSubjectName vSubject
		FROM vwSTRStudentSubjectScore
		WHERE iSubjectGradeProfileId = @iSubjectGradeProfileId 
		GROUP BY iStudentProfileId
		ORDER BY SUM(fRawScore) DESC
  END
ELSE
  BEGIN
	SELECT	iStudentProfileId ,
			SUM(fRawScore) fRawScore ,
			SUM(fGraceScore) fBSLScore ,
			@vSubjectName vSubject
		FROM vwSTRStudentSubjectScore
		WHERE iSubjectGradeProfileId in (SELECT iSubjectGradeProfileId FROM tbSTRSeisekiSpecialSubjectProfile WHERE iSeisekiIchiranId=@iSeisekiIchiranId)
		GROUP BY iStudentProfileId
		ORDER BY SUM(fRawScore) DESC
  END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

/*
 * プリントキューを監視する。
 *
 */
CREATE  PROCEDURE uspSTRWatchReport	@iPrinterId	int
AS
 DECLARE 	@iModuleReportId	int ,
		@iReportId int ,
		@vParams	varchar(1024)
 SET @iReportId=NULL
 SELECT @iReportId=iReportId,@iModuleReportId=iModuleReportId,@vParams=vParameterString
	FROM tbSTRReportData
	WHERE iStatus=0 AND iPrinterId=@iPrinterId
	ORDER BY iPriority,dtRequest
 IF ( @iReportId IS NOT NULL )
 BEGIN
	UPDATE tbSTRReportData
		SET iStatus=1,dtOutput=getdate()
		WHERE iReportId=@iReportId
 END
 SELECT 	@iModuleReportId as iModuleReportId ,
		@vParams as vParameterString

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

