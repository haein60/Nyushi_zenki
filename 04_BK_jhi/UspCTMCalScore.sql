CREATE  PROCEDURE UspCTMCalScore1 (
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
