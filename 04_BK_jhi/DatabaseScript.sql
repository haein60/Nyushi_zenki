if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEExamineeProfile_tbSTEAdmissionType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEExamineeProfile] DROP CONSTRAINT FK_tbSTEExamineeProfile_tbSTEAdmissionType
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEExamineeProfile_tbSTEAdmissionType1]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEExamineeProfile] DROP CONSTRAINT FK_tbSTEExamineeProfile_tbSTEAdmissionType1
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEExamineeRoomProfile_tbSTEExamineeProfile]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEExamineeRoomProfile] DROP CONSTRAINT FK_tbSTEExamineeRoomProfile_tbSTEExamineeProfile
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEScoreProfile_tbSTEExamineeProfile]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEScoreProfile] DROP CONSTRAINT FK_tbSTEScoreProfile_tbSTEExamineeProfile
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEExamineeProfile_tbSTEHighSchoolType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEExamineeProfile] DROP CONSTRAINT FK_tbSTEExamineeProfile_tbSTEHighSchoolType
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEInterviewRoomProfile_tbSTEInterviewerProfile]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEInterviewRoomProfile] DROP CONSTRAINT FK_tbSTEInterviewRoomProfile_tbSTEInterviewerProfile
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEExamineeProfile_tbSTELookUpTable]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEExamineeProfile] DROP CONSTRAINT FK_tbSTEExamineeProfile_tbSTELookUpTable
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEExamineeProfile_tbSTELookUpTable1]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEExamineeProfile] DROP CONSTRAINT FK_tbSTEExamineeProfile_tbSTELookUpTable1
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEExamineeProfile_tbSTELookUpTable2]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEExamineeProfile] DROP CONSTRAINT FK_tbSTEExamineeProfile_tbSTELookUpTable2
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEExamineeProfile_tbSTELookUpTable3]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEExamineeProfile] DROP CONSTRAINT FK_tbSTEExamineeProfile_tbSTELookUpTable3
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEExamineeProfile_tbSTELookUpTable4]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEExamineeProfile] DROP CONSTRAINT FK_tbSTEExamineeProfile_tbSTELookUpTable4
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEExamineeProfile_tbSTELookUpTable5]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEExamineeProfile] DROP CONSTRAINT FK_tbSTEExamineeProfile_tbSTELookUpTable5
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEExamineeProfile_tbSTELookUpTable6]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEExamineeProfile] DROP CONSTRAINT FK_tbSTEExamineeProfile_tbSTELookUpTable6
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEExamineeProfile_tbSTERoomProfile]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEExamineeProfile] DROP CONSTRAINT FK_tbSTEExamineeProfile_tbSTERoomProfile
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEExamineeRoomProfile_tbSTERoomProfile]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEExamineeRoomProfile] DROP CONSTRAINT FK_tbSTEExamineeRoomProfile_tbSTERoomProfile
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEInterviewRoomProfile_tbSTERoomProfile]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEInterviewRoomProfile] DROP CONSTRAINT FK_tbSTEInterviewRoomProfile_tbSTERoomProfile
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEScoreDetail_tbSTEScoreProfile]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEScoreDetail] DROP CONSTRAINT FK_tbSTEScoreDetail_tbSTEScoreProfile
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEScoreProfile_tbSTESubjectProfile]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEScoreProfile] DROP CONSTRAINT FK_tbSTEScoreProfile_tbSTESubjectProfile
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTESubjectQuestionProfile_tbSTESubjectProfile]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTESubjectQuestionProfile] DROP CONSTRAINT FK_tbSTESubjectQuestionProfile_tbSTESubjectProfile
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEScoreDetail_tbSTESubjectQuestionProfile]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEScoreDetail] DROP CONSTRAINT FK_tbSTEScoreDetail_tbSTESubjectQuestionProfile
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTESecondExamProfile_tbSTESystemProfile]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTESecondExamProfile] DROP CONSTRAINT FK_tbSTESecondExamProfile_tbSTESystemProfile
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbSTEExamineeProfile_tbSTEZipCodeMaster]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbSTEExamineeProfile] DROP CONSTRAINT FK_tbSTEExamineeProfile_tbSTEZipCodeMaster
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

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UspCTMAllocateRoom_org]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[UspCTMAllocateRoom_org]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UspCTMCalScore]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[UspCTMCalScore]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UspCTMTest1710]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[UspCTMTest1710]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UspCTMTest1810]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[UspCTMTest1810]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UspCTMTest1910]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[UspCTMTest1910]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[test]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[test]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbSTEAdmissionType]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbSTEAdmissionType]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbSTEExamineeProfile]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbSTEExamineeProfile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbSTEExamineeRoomProfile]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbSTEExamineeRoomProfile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbSTEHighSchoolType]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbSTEHighSchoolType]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbSTEInterviewGroupProfile]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbSTEInterviewGroupProfile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbSTEInterviewRoomProfile]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbSTEInterviewRoomProfile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbSTEInterviewerProfile]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbSTEInterviewerProfile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbSTELookUpTable]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbSTELookUpTable]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbSTERoomProfile]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbSTERoomProfile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbSTEScoreDetail]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbSTEScoreDetail]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbSTEScoreProfile]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbSTEScoreProfile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbSTESecondExamProfile]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbSTESecondExamProfile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbSTESubjectProfile]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbSTESubjectProfile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbSTESubjectQuestionProfile]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbSTESubjectQuestionProfile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbSTESystemProfile]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbSTESystemProfile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbSTETableIdMapping]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbSTETableIdMapping]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbSTEZipCodeMaster]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbSTEZipCodeMaster]
GO

CREATE TABLE [dbo].[tbSTEAdmissionType] (
	[iAdmissionType] [int] NOT NULL ,
	[vAdmissionName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dtCreate] [datetime] NULL ,
	[dtUpdate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbSTEExamineeProfile] (
	[iExamineeProfileId] [int] NOT NULL ,
	[iJukenNumber] [int] NOT NULL ,
	[iNendo] [int] NOT NULL ,
	[vExamineeName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vKanaName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vAddress] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[iZipCodeId] [int] NULL ,
	[iSex] [int] NULL ,
	[iHighSchoolId] [int] NULL ,
	[vTelephone] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vEmailAddress] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dtBirthDay] [datetime] NULL ,
	[iRoomProfileId] [int] NULL ,
	[iAbsentFlag] [int] NULL ,
	[iRejectFlag] [int] NULL ,
	[iExamineeStatus] [int] NULL ,
	[dtSecondExamDay] [datetime] NULL ,
	[iUniversityType] [int] NULL ,
	[iBackgroundId] [int] NULL ,
	[iFamilyId] [int] NULL ,
	[iParentJobCategory] [int] NULL ,
	[iQualificationId] [int] NULL ,
	[iSuisenFlagId] [int] NULL ,
	[vNationality] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[iPhysicalConditionId] [int] NULL ,
	[iLanguageSubjProfileId] [int] NULL ,
	[iScienceSubjProfileId1] [int] NULL ,
	[iScienceSubjProfileId2] [int] NULL ,
	[iPreferenceDay1Flag] [int] NULL ,
	[iPreferenceDay2Flag] [int] NULL ,
	[iPreferenceDay3Flag] [int] NULL ,
	[iMultipleApplyFlag] [int] NULL ,
	[iAdmissionType1] [int] NULL ,
	[iAdmissionType2] [int] NULL ,
	[dtCreate] [datetime] NULL ,
	[dtUpdate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbSTEExamineeRoomProfile] (
	[iExamineeRoomProfileId] [int] NOT NULL ,
	[iExamineeProfileId] [int] NULL ,
	[iRoomProfileId] [int] NULL ,
	[iSubjectProfileId] [int] NULL ,
	[dtCreate] [datetime] NULL ,
	[dtUpdate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbSTEHighSchoolType] (
	[iHighSchoolId] [int] NOT NULL ,
	[vHighSchoolCode] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vHighSchoolName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[iHighSchoolRecommendation] [int] NULL ,
	[iZipCodeId] [int] NULL ,
	[vAddress1] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vAddress2] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vTelephoneNo] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vFaxNo] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vRepresentativeName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[iLetterFlag] [int] NULL ,
	[iHighSchoolRecommendationYear1] [int] NULL ,
	[iHighSchoolRecommendationYear2] [int] NULL ,
	[iHighSchoolDropRecommendationYear] [int] NULL ,
	[dtCreate] [datetime] NULL ,
	[dtUpdate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbSTEInterviewGroupProfile] (
	[iInterviewGroupProfileId] [int] NOT NULL ,
	[vInterviewGroupName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dtCreate] [datetime] NULL ,
	[dtUpdate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbSTEInterviewRoomProfile] (
	[iInterviewRoomProfileId] [int] NOT NULL ,
	[iInterviewerProfileId] [int] NULL ,
	[iRoomProfileId] [int] NULL ,
	[iSubjectProfileId] [int] NULL ,
	[iDayFlag] [int] NULL ,
	[dtCreate] [datetime] NULL ,
	[dtUpdate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbSTEInterviewerProfile] (
	[iInterviewerProfileId] [int] NOT NULL ,
	[iInterviewGroupProfileId] [int] NULL ,
	[vInterviewerName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dtCreate] [datetime] NULL ,
	[dtUpdate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbSTELookUpTable] (
	[iLookUpTableID] [int] NOT NULL ,
	[iLookUpTableType] [int] NULL ,
	[iValue] [int] NULL ,
	[vName] [nvarchar] (64) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbSTERoomProfile] (
	[iRoomProfileId] [int] NOT NULL ,
	[vRoomName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[iRandom] [int] NULL ,
	[iMaxCapacity] [int] NULL ,
	[dtCreate] [datetime] NULL ,
	[dtUpdate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbSTEScoreDetail] (
	[iScoreDetailId] [int] NOT NULL ,
	[iScoreProfileId] [int] NULL ,
	[iSubjectQuestionId] [int] NULL ,
	[fDetailScore] [float] NULL ,
	[dtCreate] [datetime] NULL ,
	[dtUpdate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbSTEScoreProfile] (
	[iScoreProfileId] [int] NOT NULL ,
	[iSubjectProfileId] [int] NULL ,
	[iExamineeProfileId] [int] NULL ,
	[fRawScore] [float] NULL ,
	[fChoseiScore] [float] NULL ,
	[iAbsentFlag] [int] NULL ,
	[dtCreate] [datetime] NULL ,
	[dtUpdate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbSTESecondExamProfile] (
	[iSecondExamProfileId] [int] NOT NULL ,
	[iSystemProfileId] [int] NULL ,
	[dtSecondExamDay1] [datetime] NULL ,
	[dtSecondExamDay2] [datetime] NULL ,
	[dtSecondExamDay3] [datetime] NULL ,
	[iNumberOfExamineeDay1] [int] NULL ,
	[iNumberOfExamineeDay2] [int] NULL ,
	[iNumberOfExamineeDay3] [int] NULL ,
	[iNumberOfRoomDay1] [int] NULL ,
	[iNumberOfRoomDay2] [int] NULL ,
	[iNumberOfRoomDay3] [int] NULL ,
	[dtCreate] [datetime] NULL ,
	[dtUpdate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbSTESubjectProfile] (
	[iSubjectProfileId] [int] NOT NULL ,
	[vSubjectName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[iExamType] [int] NULL ,
	[dtCreate] [datetime] NULL ,
	[dtUpdate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbSTESubjectQuestionProfile] (
	[iSubjectQuestionId] [int] NOT NULL ,
	[iSubjectProfileId] [int] NULL ,
	[iQuestionNo] [int] NULL ,
	[vQuestionName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[iInterviewerProfileId] [int] NULL ,
	[dtCreate] [datetime] NULL ,
	[dtUpdate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbSTESystemProfile] (
	[iSystemProfileId] [int] NOT NULL ,
	[iNendo] [int] NULL ,
	[dtExamDate] [datetime] NULL ,
	[iCurrentPhase] [int] NULL ,
	[iVisibleFlag] [int] NULL ,
	[iActiveFlag] [int] NULL ,
	[vServerName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vSourceName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vLoginName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbSTETableIdMapping] (
	[vTableName] [char] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[iTableCounterIdMapping] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbSTEZipCodeMaster] (
	[iZipCodeId] [int] NOT NULL ,
	[vZipCodeName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vPrefectureName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vCityName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vAddress1] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vAddress2] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dtCreate] [datetime] NULL ,
	[dtUpdate] [datetime] NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tbSTEAdmissionType] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbSTEAdmissionType] PRIMARY KEY  CLUSTERED 
	(
		[iAdmissionType]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbSTEExamineeProfile] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbSTEExamineeProfile] PRIMARY KEY  CLUSTERED 
	(
		[iExamineeProfileId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbSTEExamineeRoomProfile] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbSTEExamineeRoomProfile] PRIMARY KEY  CLUSTERED 
	(
		[iExamineeRoomProfileId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbSTEHighSchoolType] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbSTEHighSchoolType] PRIMARY KEY  CLUSTERED 
	(
		[iHighSchoolId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbSTEInterviewRoomProfile] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbSTEInterviewRoomProfile] PRIMARY KEY  CLUSTERED 
	(
		[iInterviewRoomProfileId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbSTEInterviewerProfile] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbSTEInterviewerProfile] PRIMARY KEY  CLUSTERED 
	(
		[iInterviewerProfileId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbSTELookUpTable] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbSTELookUpTable] PRIMARY KEY  CLUSTERED 
	(
		[iLookUpTableID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbSTERoomProfile] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbSTERoomProfile] PRIMARY KEY  CLUSTERED 
	(
		[iRoomProfileId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbSTEScoreDetail] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbSTEScoreDetail] PRIMARY KEY  CLUSTERED 
	(
		[iScoreDetailId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbSTEScoreProfile] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbSTEScoreProfile] PRIMARY KEY  CLUSTERED 
	(
		[iScoreProfileId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbSTESecondExamProfile] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbSTESecondExamProfile] PRIMARY KEY  CLUSTERED 
	(
		[iSecondExamProfileId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbSTESubjectProfile] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbSTESubjectProfile] PRIMARY KEY  CLUSTERED 
	(
		[iSubjectProfileId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbSTESubjectQuestionProfile] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbSTESubjectQuestionProfile] PRIMARY KEY  CLUSTERED 
	(
		[iSubjectQuestionId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbSTESystemProfile] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbSTESystemProfile] PRIMARY KEY  CLUSTERED 
	(
		[iSystemProfileId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbSTETableIdMapping] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbSTETableIdMapping] PRIMARY KEY  CLUSTERED 
	(
		[vTableName]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbSTEZipCodeMaster] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbSTEZipCodeMaster] PRIMARY KEY  CLUSTERED 
	(
		[iZipCodeId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbSTEExamineeProfile] WITH NOCHECK ADD 
	CONSTRAINT [IX_tbSTEExamineeProfile_1] UNIQUE  NONCLUSTERED 
	(
		[iJukenNumber],
		[iNendo]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbSTEInterviewGroupProfile] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbSTEInterviewGroup] PRIMARY KEY  NONCLUSTERED 
	(
		[iInterviewGroupProfileId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbSTEScoreProfile] WITH NOCHECK ADD 
	CONSTRAINT [DF_tbSTEScoreProfile_dtCreate] DEFAULT (getdate()) FOR [dtCreate],
	CONSTRAINT [DF_tbSTEScoreProfile_dtUpdate] DEFAULT (getdate()) FOR [dtUpdate]
GO

 CREATE  INDEX [IX_tbSTEExamineeProfile] ON [dbo].[tbSTEExamineeProfile]([iExamineeProfileId]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [UK_RandomNo] ON [dbo].[tbSTERoomProfile]([iRandom]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tbSTEExamineeProfile] ADD 
	CONSTRAINT [FK_tbSTEExamineeProfile_tbSTEAdmissionType] FOREIGN KEY 
	(
		[iAdmissionType1]
	) REFERENCES [dbo].[tbSTEAdmissionType] (
		[iAdmissionType]
	),
	CONSTRAINT [FK_tbSTEExamineeProfile_tbSTEAdmissionType1] FOREIGN KEY 
	(
		[iAdmissionType2]
	) REFERENCES [dbo].[tbSTEAdmissionType] (
		[iAdmissionType]
	),
	CONSTRAINT [FK_tbSTEExamineeProfile_tbSTEHighSchoolType] FOREIGN KEY 
	(
		[iHighSchoolId]
	) REFERENCES [dbo].[tbSTEHighSchoolType] (
		[iHighSchoolId]
	),
	CONSTRAINT [FK_tbSTEExamineeProfile_tbSTELookUpTable] FOREIGN KEY 
	(
		[iBackgroundId]
	) REFERENCES [dbo].[tbSTELookUpTable] (
		[iLookUpTableID]
	),
	CONSTRAINT [FK_tbSTEExamineeProfile_tbSTELookUpTable1] FOREIGN KEY 
	(
		[iFamilyId]
	) REFERENCES [dbo].[tbSTELookUpTable] (
		[iLookUpTableID]
	),
	CONSTRAINT [FK_tbSTEExamineeProfile_tbSTELookUpTable2] FOREIGN KEY 
	(
		[iQualificationId]
	) REFERENCES [dbo].[tbSTELookUpTable] (
		[iLookUpTableID]
	),
	CONSTRAINT [FK_tbSTEExamineeProfile_tbSTELookUpTable3] FOREIGN KEY 
	(
		[iPhysicalConditionId]
	) REFERENCES [dbo].[tbSTELookUpTable] (
		[iLookUpTableID]
	),
	CONSTRAINT [FK_tbSTEExamineeProfile_tbSTELookUpTable4] FOREIGN KEY 
	(
		[iLanguageSubjProfileId]
	) REFERENCES [dbo].[tbSTELookUpTable] (
		[iLookUpTableID]
	),
	CONSTRAINT [FK_tbSTEExamineeProfile_tbSTELookUpTable5] FOREIGN KEY 
	(
		[iScienceSubjProfileId1]
	) REFERENCES [dbo].[tbSTELookUpTable] (
		[iLookUpTableID]
	),
	CONSTRAINT [FK_tbSTEExamineeProfile_tbSTELookUpTable6] FOREIGN KEY 
	(
		[iScienceSubjProfileId2]
	) REFERENCES [dbo].[tbSTELookUpTable] (
		[iLookUpTableID]
	),
	CONSTRAINT [FK_tbSTEExamineeProfile_tbSTERoomProfile] FOREIGN KEY 
	(
		[iRoomProfileId]
	) REFERENCES [dbo].[tbSTERoomProfile] (
		[iRoomProfileId]
	),
	CONSTRAINT [FK_tbSTEExamineeProfile_tbSTEZipCodeMaster] FOREIGN KEY 
	(
		[iZipCodeId]
	) REFERENCES [dbo].[tbSTEZipCodeMaster] (
		[iZipCodeId]
	)
GO

ALTER TABLE [dbo].[tbSTEExamineeRoomProfile] ADD 
	CONSTRAINT [FK_tbSTEExamineeRoomProfile_tbSTEExamineeProfile] FOREIGN KEY 
	(
		[iExamineeProfileId]
	) REFERENCES [dbo].[tbSTEExamineeProfile] (
		[iExamineeProfileId]
	),
	CONSTRAINT [FK_tbSTEExamineeRoomProfile_tbSTERoomProfile] FOREIGN KEY 
	(
		[iRoomProfileId]
	) REFERENCES [dbo].[tbSTERoomProfile] (
		[iRoomProfileId]
	)
GO

ALTER TABLE [dbo].[tbSTEInterviewRoomProfile] ADD 
	CONSTRAINT [FK_tbSTEInterviewRoomProfile_tbSTEInterviewerProfile] FOREIGN KEY 
	(
		[iInterviewerProfileId]
	) REFERENCES [dbo].[tbSTEInterviewerProfile] (
		[iInterviewerProfileId]
	),
	CONSTRAINT [FK_tbSTEInterviewRoomProfile_tbSTERoomProfile] FOREIGN KEY 
	(
		[iRoomProfileId]
	) REFERENCES [dbo].[tbSTERoomProfile] (
		[iRoomProfileId]
	)
GO

ALTER TABLE [dbo].[tbSTEScoreDetail] ADD 
	CONSTRAINT [FK_tbSTEScoreDetail_tbSTEScoreProfile] FOREIGN KEY 
	(
		[iScoreProfileId]
	) REFERENCES [dbo].[tbSTEScoreProfile] (
		[iScoreProfileId]
	),
	CONSTRAINT [FK_tbSTEScoreDetail_tbSTESubjectQuestionProfile] FOREIGN KEY 
	(
		[iSubjectQuestionId]
	) REFERENCES [dbo].[tbSTESubjectQuestionProfile] (
		[iSubjectQuestionId]
	)
GO

ALTER TABLE [dbo].[tbSTEScoreProfile] ADD 
	CONSTRAINT [FK_tbSTEScoreProfile_tbSTEExamineeProfile] FOREIGN KEY 
	(
		[iExamineeProfileId]
	) REFERENCES [dbo].[tbSTEExamineeProfile] (
		[iExamineeProfileId]
	),
	CONSTRAINT [FK_tbSTEScoreProfile_tbSTESubjectProfile] FOREIGN KEY 
	(
		[iSubjectProfileId]
	) REFERENCES [dbo].[tbSTESubjectProfile] (
		[iSubjectProfileId]
	)
GO

ALTER TABLE [dbo].[tbSTESecondExamProfile] ADD 
	CONSTRAINT [FK_tbSTESecondExamProfile_tbSTESystemProfile] FOREIGN KEY 
	(
		[iSystemProfileId]
	) REFERENCES [dbo].[tbSTESystemProfile] (
		[iSystemProfileId]
	)
GO

ALTER TABLE [dbo].[tbSTESubjectQuestionProfile] ADD 
	CONSTRAINT [FK_tbSTESubjectQuestionProfile_tbSTESubjectProfile] FOREIGN KEY 
	(
		[iSubjectProfileId]
	) REFERENCES [dbo].[tbSTESubjectProfile] (
		[iSubjectProfileId]
	)
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- ****************************************************************************************************************************
-- Stored Procedure		: UspCTMAllocateDay1Room
-- Input Parametrs		: Nil
-- Created			: 22/10/2001		
-- Author			: Dileep Cherian
--Output				: Nil
--Modification History		: Nil
-- Reference			: Functional Spec of Distribution Of Examinee.doc (ver 1.0)
-- ****************************************************************************************************************************
CREATE PROCEDURE UspCTMAllocateDay1Room
AS
BEGIN
DECLARE
@iExamineeId INTEGER,
@iSex INTEGER,
@iHighSchoolId INTEGER,
@iTotalExamineeDay1 INTEGER,
@iTotalRoomsDay1 INTEGER,
@Capacity INTEGER,
@Id INTEGER,
@counter INTEGER,
@iTotalMalesDay1 INTEGER,
@iTotalFemalesDay1 INTEGER,
@SchoolId INTEGER,
@iRoomId INTEGER,
@dtExamDay1 DATETIME,
@iCapacity INTEGER,
@iCount INTEGER,
@iCounter INTEGER,
@iNewId INTEGER,
@iCheckExisting INTEGER,
@iNormalInterview INTEGER,
@iNormalReport INTEGER;
CREATE TABLE #RoomDetail
(
	iRoomId INTEGER,
	iRoomProfileId INTEGER,
	iMaxCapacity INTEGER
)
CREATE TABLE #TempRoom1
(
	iRoomId INTEGER,
	iExamineeId INTEGER,
	iSex INTEGER,	
	iHighSchoolId INTEGER
)
SET @iTotalExamineeDay1 = (SELECT COUNT(*) FROM #TempDay1)
SET @iTotalRoomsDay1 = (SELECT iNoRoomDay1 FROM #TempSecondExam)
SET @iTotalMalesDay1 = (SELECT COUNT(*) FROM #TempDay1 WHERE iSex = 0)
SET @iTotalFemalesDay1 = (SELECT COUNT(*) FROM #TempDay1 WHERE iSex = 1)
SET @counter = 1
DECLARE temp_cursor2 CURSOR FOR
SELECT iRoomProfileId, iMaxCapacity FROM tbSTERoomProfile ORDER BY iRoomProfileId
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
SET @iRoomId = 1
DECLARE temp_cursor3 CURSOR FOR
SELECT DISTINCT iHighSchoolId FROM #TempDay1
OPEN temp_cursor3
FETCH NEXT FROM temp_cursor3 INTO @SchoolId
WHILE @@FETCH_STATUS = 0 
BEGIN		
	DECLARE temp_cursor4 CURSOR FOR
	SELECT iExamineeProfileId, iSex, iHighSchoolId FROM #TempDay1 WHERE iSex = 0 AND iHighSchoolId = @SchoolId
	OPEN temp_cursor4
	FETCH NEXT FROM temp_cursor4 INTO @iExamineeId, @iSex, @iHighSchoolId
	WHILE @@FETCH_STATUS = 0 
	BEGIN		
		SET @iCounter = 1
		WHILE @iCounter <= @iTotalRoomsDay1
		BEGIN
			SELECT @iCount = (SELECT COUNT(*) FROM #TempRoom1 WHERE iRoomId = @iRoomId)
			SELECT @iCapacity = (SELECT iMaxCapacity FROM #RoomDetail WHERE iRoomId = @iRoomId)
			IF @iCount < @iCapacity
			BEGIN	
				INSERT INTO #TempRoom1	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsDay1
					SET @iRoomId = 1
				BREAK;
			END
			ELSE
				SET @iRoomId = @iRoomId + 1
			SET @iCounter = @iCounter + 1
		END
		FETCH NEXT FROM temp_cursor4 INTO @iExamineeId, @iSex, @iHighSchoolId
	END
	CLOSE temp_cursor4
	DEALLOCATE temp_cursor4 
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
			IF @iCount < @iCapacity
			BEGIN	
				INSERT INTO #TempRoom1	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsDay1
					SET @iRoomId = 1
				BREAK;
			END
			ELSE
				SET @iRoomId = @iRoomId + 1
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
DECLARE temp_cursor6 CURSOR FOR
SELECT iRoomId, iExamineeId FROM #TempRoom1
OPEN temp_cursor6
FETCH NEXT FROM temp_cursor6 INTO @iRoomId, @iExamineeId
WHILE @@FETCH_STATUS = 0 
BEGIN
	/*BEGIN TRANSACTION*/
	SELECT @Id = (SELECT iRoomProfileId FROM #RoomDetail WHERE iRoomId = @iRoomId)
	SELECT @dtExamDay1 = (SELECT dtDay1 FROM #TempSecondExam)
	
	UPDATE tbSTEExamineeProfile SET dtSecondExamDay = @dtExamDay1 
	WHERE iExamineeProfileId = @iExamineeId
	SELECT @iCheckExisting = (SELECT COUNT(*) FROM tbSTEExamineeRoomProfile
	WHERE iExamineeProfileId = @iExamineeId)
	
	IF @iCheckExisting = 0 
	BEGIN
		SELECT @iNewId = (SELECT MAX(iExamineeRoomProfileId) FROM tbSTEExamineeRoomProfile)
		IF @iNewId IS NULL
			SELECT @iNewId = (SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName = 'tbSTEExamineeRoomProfile')		
		ELSE
			SET @iNewId = @iNewId + 1
		SELECT @iNormalInterview = (SELECT iSubjectProfileId FROM tbSTESubjectProfile WHERE iExamType = 2)
		INSERT INTO tbSTEExamineeRoomProfile VALUES(@iNewId, @iExamineeId, @Id, @iNormalInterview, getdate(),getdate())
		--SET @iNewId = @iNewId + 1
		--SELECT @iNormalReport = (SELECT iSubjectProfileId FROM tbSTESubjectProfile WHERE iExamType = 3)
		--INSERT INTO tbSTEExamineeRoomProfile VALUES(@iNewId, @iExamineeId, @Id, @iNormalReport, getdate(),getdate())
	END
	/*END TRANSACTION*/
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
@iTotalExamineeDay2 INTEGER,
@iTotalRoomsDay2 INTEGER,
@Capacity INTEGER,
@Id INTEGER,
@counter INTEGER,
@iTotalMalesDay2 INTEGER,
@iTotalFemalesDay2 INTEGER,
@SchoolId INTEGER,
@iRoomId INTEGER,
@dtExamDay2 DATETIME,
@iCapacity INTEGER,
@iCount INTEGER,
@iCounter INTEGER,
@iNewId INTEGER,
@iCheckExisting INTEGER,
@iNormalInterview INTEGER,
@iNormalReport INTEGER;
CREATE TABLE #RoomDetail
(
	iRoomId INTEGER,
	iRoomProfileId INTEGER,
	iMaxCapacity INTEGER
)
CREATE TABLE #TempRoom2
(
	iRoomId INTEGER,
	iExamineeId INTEGER,
	iSex INTEGER,	
	iHighSchoolId INTEGER
)
SET @iTotalExamineeDay2 = (SELECT COUNT(*) FROM #TempDay2)
SET @iTotalRoomsDay2 = (SELECT iNoRoomDay2 FROM #TempSecondExam)
SET @iTotalMalesDay2 = (SELECT COUNT(*) FROM #TempDay2 WHERE iSex = 0)
SET @iTotalFemalesDay2 = (SELECT COUNT(*) FROM #TempDay2 WHERE iSex = 1)
SET @counter = 1
DECLARE temp_cursor2 CURSOR FOR
SELECT iRoomProfileId, iMaxCapacity FROM tbSTERoomProfile ORDER BY iRoomProfileId
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
SET @iRoomId = 1
DECLARE temp_cursor3 CURSOR FOR
SELECT DISTINCT iHighSchoolId FROM #TempDay2
OPEN temp_cursor3
FETCH NEXT FROM temp_cursor3 INTO @SchoolId
WHILE @@FETCH_STATUS = 0 
BEGIN
	DECLARE temp_cursor4 CURSOR FOR
	SELECT iExamineeProfileId, iSex, iHighSchoolId FROM #TempDay2 WHERE iSex = 0 AND iHighSchoolId = @SchoolId
	OPEN temp_cursor4
	FETCH NEXT FROM temp_cursor4 INTO @iExamineeId, @iSex, @iHighSchoolId
	WHILE @@FETCH_STATUS = 0 
	BEGIN
		SET @iCounter = 1
		WHILE @iCounter <= @iTotalRoomsDay2
		BEGIN
			SELECT @iCount = (SELECT COUNT(*) FROM #TempRoom2 WHERE iRoomId = @iRoomId)
			SELECT @iCapacity = (SELECT iMaxCapacity FROM #RoomDetail WHERE iRoomId = @iRoomId)
			IF @iCount < @iCapacity
			BEGIN	
				INSERT INTO #TempRoom2	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsDay2
					SET @iRoomId = 1
				BREAK;
			END
			ELSE
				SET @iRoomId = @iRoomId + 1
			SET @iCounter = @iCounter + 1
		END
		FETCH NEXT FROM temp_cursor4 INTO @iExamineeId, @iSex, @iHighSchoolId
	END
	CLOSE temp_cursor4
	DEALLOCATE temp_cursor4 
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
			IF @iCount < @iCapacity
			BEGIN	
				INSERT INTO #TempRoom2	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsDay2
					SET @iRoomId = 1
				BREAK;
			END
			ELSE
				SET @iRoomId = @iRoomId + 1
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
DECLARE temp_cursor6 CURSOR FOR
SELECT iRoomId, iExamineeId, iSex, iHighSchoolId FROM #TempRoom2
OPEN temp_cursor6
FETCH NEXT FROM temp_cursor6 INTO @iRoomId, @iExamineeId, @iSex, @iHighSchoolId
WHILE @@FETCH_STATUS = 0 
BEGIN
	SELECT @Id = (SELECT iRoomProfileId FROM #RoomDetail WHERE iRoomId = @iRoomId)
	SELECT @dtExamDay2 = (SELECT dtDay2 FROM #TempSecondExam)
	UPDATE tbSTEExamineeProfile SET dtSecondExamDay = @dtExamDay2 
	WHERE iExamineeProfileId = @iExamineeId
	
	SELECT @iCheckExisting = (SELECT COUNT(*) FROM tbSTEExamineeRoomProfile
	WHERE iExamineeProfileId = @iExamineeId)
	
	IF @iCheckExisting = 0 
	BEGIN
		SELECT @iNewId = (SELECT MAX(iExamineeRoomProfileId) FROM tbSTEExamineeRoomProfile)
		IF @iNewId IS NULL
		BEGIN
			SELECT @iNewId = (SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName = 'tbSTEExamineeRoomProfile')
		END
		ELSE
			SET @iNewId = @iNewId + 1
		SELECT @iNormalInterview = (SELECT iSubjectProfileId FROM tbSTESubjectProfile WHERE iExamType = 2)
		INSERT INTO tbSTEExamineeRoomProfile VALUES(@iNewId, @iExamineeId, @Id, @iNormalInterview, getdate(),getdate())
		--SET @iNewId = @iNewId + 1
		--SELECT @iNormalReport = (SELECT iSubjectProfileId FROM tbSTESubjectProfile WHERE iExamType = 3)
		--INSERT INTO tbSTEExamineeRoomProfile VALUES(@iNewId, @iExamineeId, @Id, @iNormalReport, getdate(),getdate())
	END
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
SET ANSI_NULLS ON 
GO

-- ****************************************************************************************************************************
-- Stored Procedure		: UspCTMAllocateDay3Room
-- Input Parametrs		: Nil
-- Created			: 22/10/2001		
-- Author			: Dileep Cherian
--Output				: Nil
--Modification History		: Nil
-- Reference			: Functional Spec of Distribution Of Examinee.doc (ver 1.0)
-- ****************************************************************************************************************************
CREATE PROCEDURE UspCTMAllocateDay3Room
AS
BEGIN
DECLARE
@iExamineeId INTEGER,
@iSex INTEGER,
@iHighSchoolId INTEGER,
@iTotalExamineeDay3 INTEGER,
@iTotalRoomsDay3 INTEGER,
@Capacity INTEGER,
@Id INTEGER,
@counter INTEGER,
@iTotalMalesDay3 INTEGER,
@iTotalFemalesDay3 INTEGER,
@SchoolId INTEGER,
@iRoomId INTEGER,
@dtExamDay3 DATETIME,
@iCapacity INTEGER,
@iCount INTEGER,
@iCounter INTEGER,
@iNewId INTEGER,
@iCheckExisting INTEGER,
@iNormalInterview INTEGER,
@iNormalReport INTEGER;
CREATE TABLE #RoomDetail
(
	iRoomId INTEGER,
	iRoomProfileId INTEGER,
	iMaxCapacity INTEGER
)
CREATE TABLE #TempRoom3
(
	iRoomId INTEGER,
	iExamineeId INTEGER,
	iSex INTEGER,	
	iHighSchoolId INTEGER
)
SET @iTotalExamineeDay3 = (SELECT COUNT(*) FROM #TempDay3)
SET @iTotalRoomsDay3 = (SELECT iNoRoomDay3 FROM #TempSecondExam)
SET @iTotalMalesDay3 = (SELECT COUNT(*) FROM #TempDay3 WHERE iSex = 0)
SET @iTotalFemalesDay3 = (SELECT COUNT(*) FROM #TempDay3 WHERE iSex = 1)
SET @counter = 1
DECLARE temp_cursor2 CURSOR FOR
SELECT iRoomProfileId, iMaxCapacity FROM tbSTERoomProfile ORDER BY iRoomProfileId
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
SET @iRoomId = 1
DECLARE temp_cursor3 CURSOR FOR
SELECT DISTINCT iHighSchoolId FROM #TempDay3
OPEN temp_cursor3
FETCH NEXT FROM temp_cursor3 INTO @SchoolId
WHILE @@FETCH_STATUS = 0 
BEGIN
	DECLARE temp_cursor4 CURSOR FOR
	SELECT iExamineeProfileId, iSex, iHighSchoolId FROM #TempDay3 WHERE iSex = 0 AND iHighSchoolId = @SchoolId
	OPEN temp_cursor4
	FETCH NEXT FROM temp_cursor4 INTO @iExamineeId, @iSex, @iHighSchoolId
	WHILE @@FETCH_STATUS = 0 
	BEGIN
		SET @iCounter = 1
		WHILE @iCounter <= @iTotalRoomsDay3
		BEGIN
			SELECT @iCount = (SELECT COUNT(*) FROM #TempRoom3 WHERE iRoomId = @iRoomId)
			SELECT @iCapacity = (SELECT iMaxCapacity FROM #RoomDetail WHERE iRoomId = @iRoomId)
			IF @iCount < @iCapacity
			BEGIN	
				INSERT INTO #TempRoom3	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsDay3
					SET @iRoomId = 1
				BREAK;
			END
			ELSE
				SET @iRoomId = @iRoomId + 1
			SET @iCounter = @iCounter + 1
		END
		FETCH NEXT FROM temp_cursor4 INTO @iExamineeId, @iSex, @iHighSchoolId
	END
	CLOSE temp_cursor4
	DEALLOCATE temp_cursor4 
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
			IF @iCount < @iCapacity
			BEGIN	
				INSERT INTO #TempRoom3	VALUES(	@iRoomId, @iExamineeId, @iSex, @iHighSchoolId)
				SET @iRoomId = @iRoomId + 1
				IF @iRoomId > @iTotalRoomsDay3
					SET @iRoomId = 1
				BREAK;
			END
			ELSE
				SET @iRoomId = @iRoomId + 1
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
--SELECT * FROM #TempRoom3
DECLARE temp_cursor6 CURSOR FOR
SELECT iRoomId, iExamineeId FROM #TempRoom3
OPEN temp_cursor6
FETCH NEXT FROM temp_cursor6 INTO @iRoomId, @iExamineeId
WHILE @@FETCH_STATUS = 0 
BEGIN
	SELECT @Id = (SELECT iRoomProfileId FROM #RoomDetail WHERE iRoomId = @iRoomId)
	--PRINT @Id
	SELECT @dtExamDay3 = (SELECT dtDay3 FROM #TempSecondExam)
	UPDATE tbSTEExamineeProfile SET dtSecondExamDay = @dtExamDay3 
	WHERE iExamineeProfileId = @iExamineeId
	
	SELECT @iCheckExisting = (SELECT COUNT(*) FROM tbSTEExamineeRoomProfile
	WHERE iExamineeProfileId = @iExamineeId)
	
	IF @iCheckExisting = 0 
	BEGIN
		SELECT @iNewId = (SELECT MAX(iExamineeRoomProfileId) FROM tbSTEExamineeRoomProfile)
		IF @iNewId IS NULL
			SELECT @iNewId = (SELECT iTableCounterIdMapping FROM tbSTETableIdMapping WHERE vTableName = 'tbSTEExamineeRoomProfile')
		ELSE
			SET @iNewId = @iNewId + 1
		
		SELECT @iNormalInterview = (SELECT iSubjectProfileId FROM tbSTESubjectProfile WHERE iExamType = 2)
		INSERT INTO tbSTEExamineeRoomProfile VALUES(@iNewId, @iExamineeId, @Id, @iNormalInterview, getdate(),getdate())
		
		--SET @iNewId = @iNewId + 1
		--SELECT @iNormalReport = (SELECT iSubjectProfileId FROM tbSTESubjectProfile WHERE iExamType = 3)
		--INSERT INTO tbSTEExamineeRoomProfile VALUES(@iNewId, @iExamineeId, @Id, @iNormalReport, getdate(),getdate())
	END
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
SET ANSI_NULLS ON 
GO

-- ****************************************************************************************************************************
-- Stored Procedure		: UspCTMAllocateRoom
-- Input Parametrs		: Nil
-- Created			: 22/10/2001		
-- Author			: Dileep Cherian
--Output				: Recordsets
--Modification History		: 
-- Reference			: Functional Spec of Distribution Of Examinee.doc (ver 1.0)
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
SET @iCountDay1 = 0
SET @iCountDay2 = 0
SET @iCountDay3 = 0
SET @iNoOfConditions = 1
WHILE @iNoOfConditions <= 3
BEGIN
DECLARE ExamineeCursor CURSOR FOR
SELECT iExamineeProfileId, iSex, iPreferenceDay1Flag, iPreferenceDay2Flag, iPreferenceDay3Flag, iMultipleApplyFlag, iHighSchoolId
FROM tbSTEExamineeProfile
WHERE iNendo=(SELECT iNendo FROM tbSTESystemProfile WHERE iActiveFlag=1)
AND iExamineeStatus = 1 AND iAbsentFlag = 0
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
	    SET @iCountDay1 = @iCountDay1 + 1
	    IF @iCountDay1 <= @iNoOfExamineeDay1
	        INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE
	        INSERT INTO #Day1Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
        END    
    
        IF @iDay1Flag = 0 AND @iDay2Flag = 1 AND @iDay3Flag = 0 AND @iMultipleFlag = @MultipleValue
        BEGIN
	    SET @iCountDay2 = @iCountDay2 + 1
	    IF @iCountDay2 <= @iNoOfExamineeDay2
	        INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE
	        INSERT INTO #Day2Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
        END 
        IF @iDay1Flag = 0 AND @iDay2Flag = 0 AND @iDay3Flag = 1 AND @iMultipleFlag = @MultipleValue
        BEGIN
	    SET @iCountDay3 = @iCountDay3 + 1
	    IF @iCountDay3 <= @iNoOfExamineeDay3
	        INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)	
	    ELSE
	        INSERT INTO #Day3Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
        END 
    END     
    ELSE	/* condition <> 1 or 2 */
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
    FETCH NEXT FROM ExamineeCursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
END	/* End While - Cursor */
CLOSE ExamineeCursor
DEALLOCATE ExamineeCursor
SET @iNoOfConditions = @iNoOfConditions + 1
END	/* End While - Conditions */
/* Insert the records of #TempDay12 */
DECLARE temp_cursor CURSOR FOR		
SELECT * FROM #TempDay12
OPEN temp_cursor
FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
WHILE @@FETCH_STATUS = 0
BEGIN    
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN
	    SET @iCountDay1 = @iCountDay1 - 1
	    SET @iCountDay2 = @iCountDay2 + 1
	    IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
	    	INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE	/* Insert into Day 3 */
	    BEGIN
	        SET @iCountDay2 = @iCountDay2 - 1
		     INSERT INTO #Day12Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    END
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
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN
	    SET @iCountDay1 = @iCountDay1 - 1
	    SET @iCountDay3 = @iCountDay3 + 1
	    IF @iCountDay3 <= @iNoOfExamineeDay3	/* Insert into Day 3 */
	    	INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE	/* Insert into Day 2 */
	    BEGIN
	        SET @iCountDay3 = @iCountDay3 - 1
	        INSERT INTO #Day13Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    END
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
	SET @iCountDay2 = @iCountDay2 + 1
	IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
	    INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN
	    SET @iCountDay2 = @iCountDay2 - 1
	    SET @iCountDay3 = @iCountDay3 + 1
	    IF @iCountDay3 <= @iNoOfExamineeDay3	/* Insert into Day 3 */
	    	INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE	/* Insert into Day 1 */
	    BEGIN
	        SET @iCountDay3 = @iCountDay3 - 1
			  INSERT INTO #Day23Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    END
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
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN
	    SET @iCountDay1 = @iCountDay1 - 1
	    SET @iCountDay2 = @iCountDay2 + 1
	    IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
	    	INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE	/* Insert into Day 3 */
	    BEGIN
	        SET @iCountDay2 = @iCountDay2 - 1
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
	SET @iCountDay2 = @iCountDay2 + 1
	IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
	    INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN	    
	    SET @iCountDay2 = @iCountDay2 - 1
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
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN	    
	    SET @iCountDay1 = @iCountDay1 - 1
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
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 2 */
	BEGIN	    
	    SET @iCountDay1 = @iCountDay1 - 1
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
SELECT * FROM #Day1Excess
SELECT * FROM #Day2Excess
SELECT * FROM #Day3Excess
SELECT * FROM #Day12Excess
SELECT * FROM #Day13Excess
SELECT * FROM #Day23Excess
END	/* Final End */
-- exec UspCTMAllocateRoom

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE UspCTMAllocateRoom_org
AS
BEGIN
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
SET @iCountDay1 = 0
SET @iCountDay2 = 0
SET @iCountDay3 = 0
SET @iNoOfConditions = 1
WHILE @iNoOfConditions <= 3
BEGIN
DECLARE ExamineeCursor CURSOR FOR
SELECT iExamineeProfileId, iSex, iPreferenceDay1Flag, iPreferenceDay2Flag, iPreferenceDay3Flag, iMultipleApplyFlag, iHighSchoolId
FROM tbSTEExamineeProfile
WHERE iNendo=(SELECT iNendo FROM tbSTESystemProfile WHERE iActiveFlag=1)
AND iExamineeStatus = 1 AND iAbsentFlag = 0
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
	    SET @iCountDay1 = @iCountDay1 + 1
	    IF @iCountDay1 <= @iNoOfExamineeDay1
	        INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE
	        INSERT INTO #Day1Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
        END    
    
        IF @iDay1Flag = 0 AND @iDay2Flag = 1 AND @iDay3Flag = 0 AND @iMultipleFlag = @MultipleValue
        BEGIN
	    SET @iCountDay2 = @iCountDay2 + 1
	    IF @iCountDay2 <= @iNoOfExamineeDay2
	        INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE
	        INSERT INTO #Day2Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
        END 
        IF @iDay1Flag = 0 AND @iDay2Flag = 0 AND @iDay3Flag = 1 AND @iMultipleFlag = @MultipleValue
        BEGIN
	    SET @iCountDay3 = @iCountDay3 + 1
	    IF @iCountDay3 <= @iNoOfExamineeDay3
	        INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)	
	    ELSE
	        INSERT INTO #Day3Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
        END 
    END     
    ELSE	/* condition <> 1 or 2 */
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
    FETCH NEXT FROM ExamineeCursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
END	/* End While - Cursor */
CLOSE ExamineeCursor
DEALLOCATE ExamineeCursor
SET @iNoOfConditions = @iNoOfConditions + 1
END	/* End While - Conditions */
/* Insert the records of #TempDay12 */
DECLARE temp_cursor CURSOR FOR		
SELECT * FROM #TempDay12
OPEN temp_cursor
FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
WHILE @@FETCH_STATUS = 0
BEGIN    
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN
	    SET @iCountDay1 = @iCountDay1 - 1
	    SET @iCountDay2 = @iCountDay2 + 1
	    IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
	    	INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE	/* Insert into Day 3 */
	    BEGIN
	        SET @iCountDay2 = @iCountDay2 - 1
		SET @iCountDay3 = @iCountDay3 + 1
		INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)                                
	    END
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
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN
	    SET @iCountDay1 = @iCountDay1 - 1
	    SET @iCountDay3 = @iCountDay3 + 1
	    IF @iCountDay3 <= @iNoOfExamineeDay3	/* Insert into Day 3 */
	    	INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE	/* Insert into Day 2 */
	    BEGIN
	        SET @iCountDay3 = @iCountDay3 - 1
		SET @iCountDay2 = @iCountDay2 + 1
		INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)                                
	    END
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
	SET @iCountDay2 = @iCountDay2 + 1
	IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
	    INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN
	    SET @iCountDay2 = @iCountDay2 - 1
	    SET @iCountDay3 = @iCountDay3 + 1
	    IF @iCountDay3 <= @iNoOfExamineeDay3	/* Insert into Day 3 */
	    	INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE	/* Insert into Day 1 */
	    BEGIN
	        SET @iCountDay3 = @iCountDay3 - 1
		SET @iCountDay1 = @iCountDay1 + 1
		INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)                                
	    END
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
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN
	    SET @iCountDay1 = @iCountDay1 - 1
	    SET @iCountDay2 = @iCountDay2 + 1
	    IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
	    	INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE	/* Insert into Day 3 */
	    BEGIN
	        SET @iCountDay2 = @iCountDay2 - 1
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
	SET @iCountDay2 = @iCountDay2 + 1
	IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
	    INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN	    
	    SET @iCountDay2 = @iCountDay2 - 1
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
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN	    
	    SET @iCountDay1 = @iCountDay1 - 1
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
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 2 */
	BEGIN	    
	    SET @iCountDay1 = @iCountDay1 - 1
	    SET @iCountDay2 = @iCountDay2 + 1
            INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)                                
	END
	FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
END
CLOSE temp_cursor
DEALLOCATE temp_cursor
exec UspCTMAllocateDay1Room
exec UspCTMAllocateDay2Room
exec UspCTMAllocateDay3Room
SELECT * FROM #Day1Excess
SELECT * FROM #Day2Excess
SELECT * FROM #Day3Excess
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
-- Stored Procedure		: UspCTMCalScore
-- Input Parametrs		: Nil
-- Created			: 14/09/2001		
-- Author			: Dileep Cherian
--Output				: float
--Modification History		: Nil
-- Reference			: Functional Spec of Distribution Of Examinee.doc (ver 1.0)
-- ****************************************************************************************************************************
CREATE PROCEDURE UspCTMCalScore (
@ExamType int,
@SubjectProfileId int,
@NumberOfParams int,
@Score1 float,
@Score2 float,
@Score3 float,
@Score4 float,
@Score5 float,
@Score6 float,
@Score7 float,
@Score8 float,
@Score9 float,
@Score10 float,
@TotalScore float OUTPUT
)
AS
BEGIN	
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

CREATE PROCEDURE UspCTMTest1710
AS
BEGIN
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
/* get the total no of examinees allotted for Day 1 */
SELECT @iNoOfExamineeDay1=(SELECT iNumberOfExamineeDay1 FROM tbSTESecondExamProfile
WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile
WHERE iActiveFlag=1))
/* get the total no of examinees allotted for Day 2 */
SELECT @iNoOfExamineeDay2=(SELECT iNumberOfExamineeDay2 FROM tbSTESecondExamProfile
WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile
WHERE iActiveFlag=1))
/* get the total no of examinees allotted for Day 3 */
SELECT @iNoOfExamineeDay3=(SELECT iNumberOfExamineeDay3 FROM tbSTESecondExamProfile
WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile
WHERE iActiveFlag=1))
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
SET @iCountDay1 = 0
SET @iCountDay2 = 0
SET @iCountDay3 = 0
SET @iNoOfConditions = 1
WHILE @iNoOfConditions <= 3
BEGIN
DECLARE ExamineeCursor CURSOR FOR
SELECT iExamineeProfileId, iSex, iPreferenceDay1Flag, iPreferenceDay2Flag, iPreferenceDay3Flag, iMultipleApplyFlag, iHighSchoolId
FROM tbSTEExamineeProfile
WHERE iNendo=(SELECT iNendo FROM tbSTESystemProfile WHERE iActiveFlag=1) AND iExamineeStatus = 1
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
	    SET @iCountDay1 = @iCountDay1 + 1
	    IF @iCountDay1 <= @iNoOfExamineeDay1
	        INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE
	        INSERT INTO #Day1Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
        END    
    
        IF @iDay1Flag = 0 AND @iDay2Flag = 1 AND @iDay3Flag = 0 AND @iMultipleFlag = @MultipleValue
        BEGIN
	    SET @iCountDay2 = @iCountDay2 + 1
	    IF @iCountDay2 <= @iNoOfExamineeDay2
	        INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE
	        INSERT INTO #Day2Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
        END 
        IF @iDay1Flag = 0 AND @iDay2Flag = 0 AND @iDay3Flag = 1 AND @iMultipleFlag = @MultipleValue
        BEGIN
	    SET @iCountDay3 = @iCountDay3 + 1
	    IF @iCountDay3 <= @iNoOfExamineeDay3
	        INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)	
	    ELSE
	        INSERT INTO #Day1Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
        END 
    END     
    ELSE	/* condition <> 1 or 2 */
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
    FETCH NEXT FROM ExamineeCursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
END	/* End While - Cursor */
CLOSE ExamineeCursor
DEALLOCATE ExamineeCursor
SET @iNoOfConditions = @iNoOfConditions + 1
END	/* End While - Conditions */
/* Display */
SELECT * FROM #TempDay1
SELECT * FROM #TempDay2
SELECT * FROM #TempDay3
SELECT * FROM #Day1Excess
SELECT * FROM #Day2Excess
SELECT * FROM #Day3Excess
SELECT * FROM #TempDay12
SELECT * FROM #TempDay13
SELECT * FROM #TempDay23
SELECT * FROM #TempDay123
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

CREATE PROCEDURE UspCTMTest1810
AS
BEGIN
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
CREATE TABLE #TempSecondExam
(
	iNoExamineeDay1 INTEGER,
	iNoExamineeDay2 INTEGER,
	iNoExamineeDay3 INTEGER,
	iNoRoomDay1 INTEGER,
	iNoRoomDay2 INTEGER,
	iNoRoomDay3 INTEGER
)
INSERT INTO #TempSecondExam 
SELECT iNumberOfExamineeDay1, iNumberOfExamineeDay2, iNumberOfExamineeDay3,
iNumberOfRoomDay1, iNumberOfRoomDay2, iNumberOfRoomDay3 FROM tbSTESecondExamProfile 
WHERE iSystemProfileId=(SELECT iSystemProfileId FROM tbSTESystemProfile WHERE iActiveFlag=1)
SELECT @iNoOfExamineeDay1 = (SELECT iNoExamineeDay1 FROM #TempSecondExam)
SELECT @iNoOfExamineeDay2 = (SELECT iNoExamineeDay3 FROM #TempSecondExam)
SELECT @iNoOfExamineeDay3 = (SELECT iNoExamineeDay2 FROM #TempSecondExam)
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
SET @iCountDay1 = 0
SET @iCountDay2 = 0
SET @iCountDay3 = 0
SET @iNoOfConditions = 1
WHILE @iNoOfConditions <= 3
BEGIN
DECLARE ExamineeCursor CURSOR FOR
SELECT iExamineeProfileId, iSex, iPreferenceDay1Flag, iPreferenceDay2Flag, iPreferenceDay3Flag, iMultipleApplyFlag, iHighSchoolId
FROM tbSTEExamineeProfile
WHERE iNendo=(SELECT iNendo FROM tbSTESystemProfile WHERE iActiveFlag=1) AND iExamineeStatus = 1
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
	    SET @iCountDay1 = @iCountDay1 + 1
	    IF @iCountDay1 <= @iNoOfExamineeDay1
	        INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE
	        INSERT INTO #Day1Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
        END    
    
        IF @iDay1Flag = 0 AND @iDay2Flag = 1 AND @iDay3Flag = 0 AND @iMultipleFlag = @MultipleValue
        BEGIN
	    SET @iCountDay2 = @iCountDay2 + 1
	    IF @iCountDay2 <= @iNoOfExamineeDay2
	        INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE
	        INSERT INTO #Day2Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
        END 
        IF @iDay1Flag = 0 AND @iDay2Flag = 0 AND @iDay3Flag = 1 AND @iMultipleFlag = @MultipleValue
        BEGIN
	    SET @iCountDay3 = @iCountDay3 + 1
	    IF @iCountDay3 <= @iNoOfExamineeDay3
	        INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)	
	    ELSE
	        INSERT INTO #Day1Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
        END 
    END     
    ELSE	/* condition <> 1 or 2 */
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
    FETCH NEXT FROM ExamineeCursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
END	/* End While - Cursor */
CLOSE ExamineeCursor
DEALLOCATE ExamineeCursor
SET @iNoOfConditions = @iNoOfConditions + 1
END	/* End While - Conditions */
/* Display 
SELECT * FROM #TempDay1
SELECT * FROM #TempDay2
SELECT * FROM #TempDay3
SELECT * FROM #Day1Excess
SELECT * FROM #Day2Excess
SELECT * FROM #Day3Excess
SELECT * FROM #TempDay12
SELECT * FROM #TempDay13
SELECT * FROM #TempDay23
SELECT * FROM #TempDay123
SELECT COUNT(*) 'count TempDay12' FROM #TempDay12
SELECT COUNT(*) 'count TempDay13' FROM #TempDay13
SELECT COUNT(*) 'count TempDay23' FROM #TempDay23
SELECT COUNT(*) 'count TempDay123' FROM #TempDay123*/
/* Insert the records of #TempDay12 */
DECLARE temp_cursor CURSOR FOR		
SELECT * FROM #TempDay12
OPEN temp_cursor
FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
WHILE @@FETCH_STATUS = 0
BEGIN    
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN
	    SET @iCountDay1 = @iCountDay1 - 1
	    SET @iCountDay2 = @iCountDay2 + 1
	    IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
	    	INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE	/* Insert into Day 3 */
	    BEGIN
	        SET @iCountDay2 = @iCountDay2 - 1
		SET @iCountDay3 = @iCountDay3 + 1
		INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)                                
	    END
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
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN
	    SET @iCountDay1 = @iCountDay1 - 1
	    SET @iCountDay3 = @iCountDay3 + 1
	    IF @iCountDay3 <= @iNoOfExamineeDay3	/* Insert into Day 3 */
	    	INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE	/* Insert into Day 2 */
	    BEGIN
	        SET @iCountDay3 = @iCountDay3 - 1
		SET @iCountDay2 = @iCountDay2 + 1
		INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)                                
	    END
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
	SET @iCountDay2 = @iCountDay2 + 1
	IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
	    INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN
	    SET @iCountDay2 = @iCountDay2 - 1
	    SET @iCountDay3 = @iCountDay3 + 1
	    IF @iCountDay3 <= @iNoOfExamineeDay3	/* Insert into Day 3 */
	    	INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE	/* Insert into Day 1 */
	    BEGIN
	        SET @iCountDay3 = @iCountDay3 - 1
		SET @iCountDay1 = @iCountDay1 + 1
		INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)                                
	    END
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
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN
	    SET @iCountDay1 = @iCountDay1 - 1
	    SET @iCountDay2 = @iCountDay2 + 1
	    IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
	    	INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE	/* Insert into Day 3 */
	    BEGIN
	        SET @iCountDay2 = @iCountDay2 - 1
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
	SET @iCountDay2 = @iCountDay2 + 1
	IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
	    INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN	    
	    SET @iCountDay2 = @iCountDay2 - 1
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
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN	    
	    SET @iCountDay1 = @iCountDay1 - 1
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
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 2 */
	BEGIN	    
	    SET @iCountDay1 = @iCountDay1 - 1
	    SET @iCountDay2 = @iCountDay2 + 1
            INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)                                
	END
	FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
END
CLOSE temp_cursor
DEALLOCATE temp_cursor
/* Final Display */
SELECT * FROM #TempDay1
SELECT * FROM #TempDay2
SELECT * FROM #TempDay3
exec UspCTMAllocateRoom
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

CREATE PROCEDURE UspCTMTest1910
AS
BEGIN
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
SELECT @iNoOfExamineeDay2 = (SELECT iNoExamineeDay3 FROM #TempSecondExam)
SELECT @iNoOfExamineeDay3 = (SELECT iNoExamineeDay2 FROM #TempSecondExam)
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
SET @iCountDay1 = 0
SET @iCountDay2 = 0
SET @iCountDay3 = 0
SET @iNoOfConditions = 1
WHILE @iNoOfConditions <= 3
BEGIN
DECLARE ExamineeCursor CURSOR FOR
SELECT iExamineeProfileId, iSex, iPreferenceDay1Flag, iPreferenceDay2Flag, iPreferenceDay3Flag, iMultipleApplyFlag, iHighSchoolId
FROM tbSTEExamineeProfile
WHERE iNendo=(SELECT iNendo FROM tbSTESystemProfile WHERE iActiveFlag=1) AND iExamineeStatus = 1
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
	    SET @iCountDay1 = @iCountDay1 + 1
	    IF @iCountDay1 <= @iNoOfExamineeDay1
	        INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE
	        INSERT INTO #Day1Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
        END    
    
        IF @iDay1Flag = 0 AND @iDay2Flag = 1 AND @iDay3Flag = 0 AND @iMultipleFlag = @MultipleValue
        BEGIN
	    SET @iCountDay2 = @iCountDay2 + 1
	    IF @iCountDay2 <= @iNoOfExamineeDay2
	        INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE
	        INSERT INTO #Day2Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
        END 
        IF @iDay1Flag = 0 AND @iDay2Flag = 0 AND @iDay3Flag = 1 AND @iMultipleFlag = @MultipleValue
        BEGIN
	    SET @iCountDay3 = @iCountDay3 + 1
	    IF @iCountDay3 <= @iNoOfExamineeDay3
	        INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)	
	    ELSE
	        INSERT INTO #Day1Excess VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
        END 
    END     
    ELSE	/* condition <> 1 or 2 */
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
    FETCH NEXT FROM ExamineeCursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
END	/* End While - Cursor */
CLOSE ExamineeCursor
DEALLOCATE ExamineeCursor
SET @iNoOfConditions = @iNoOfConditions + 1
END	/* End While - Conditions */
/* Insert the records of #TempDay12 */
DECLARE temp_cursor CURSOR FOR		
SELECT * FROM #TempDay12
OPEN temp_cursor
FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
WHILE @@FETCH_STATUS = 0
BEGIN    
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN
	    SET @iCountDay1 = @iCountDay1 - 1
	    SET @iCountDay2 = @iCountDay2 + 1
	    IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
	    	INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE	/* Insert into Day 3 */
	    BEGIN
	        SET @iCountDay2 = @iCountDay2 - 1
		SET @iCountDay3 = @iCountDay3 + 1
		INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)                                
	    END
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
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN
	    SET @iCountDay1 = @iCountDay1 - 1
	    SET @iCountDay3 = @iCountDay3 + 1
	    IF @iCountDay3 <= @iNoOfExamineeDay3	/* Insert into Day 3 */
	    	INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE	/* Insert into Day 2 */
	    BEGIN
	        SET @iCountDay3 = @iCountDay3 - 1
		SET @iCountDay2 = @iCountDay2 + 1
		INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)                                
	    END
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
	SET @iCountDay2 = @iCountDay2 + 1
	IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
	    INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN
	    SET @iCountDay2 = @iCountDay2 - 1
	    SET @iCountDay3 = @iCountDay3 + 1
	    IF @iCountDay3 <= @iNoOfExamineeDay3	/* Insert into Day 3 */
	    	INSERT INTO #TempDay3 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE	/* Insert into Day 1 */
	    BEGIN
	        SET @iCountDay3 = @iCountDay3 - 1
		SET @iCountDay1 = @iCountDay1 + 1
		INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)                                
	    END
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
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN
	    SET @iCountDay1 = @iCountDay1 - 1
	    SET @iCountDay2 = @iCountDay2 + 1
	    IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
	    	INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	    ELSE	/* Insert into Day 3 */
	    BEGIN
	        SET @iCountDay2 = @iCountDay2 - 1
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
	SET @iCountDay2 = @iCountDay2 + 1
	IF @iCountDay2 <= @iNoOfExamineeDay2	/* Insert into Day 2 */
	    INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN	    
	    SET @iCountDay2 = @iCountDay2 - 1
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
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 3 */
	BEGIN	    
	    SET @iCountDay1 = @iCountDay1 - 1
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
	SET @iCountDay1 = @iCountDay1 + 1
	IF @iCountDay1 <= @iNoOfExamineeDay1	/* Insert into Day 1 */
	    INSERT INTO #TempDay1 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)
	ELSE	/* Insert into Day 2 */
	BEGIN	    
	    SET @iCountDay1 = @iCountDay1 - 1
	    SET @iCountDay2 = @iCountDay2 + 1
            INSERT INTO #TempDay2 VALUES(@iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId)                                
	END
	FETCH NEXT FROM temp_cursor INTO @iExamineeId, @iSex, @iDay1Flag, @iDay2Flag, @iDay3Flag, @iMultipleFlag, @iHighSchoolId
END
CLOSE temp_cursor
DEALLOCATE temp_cursor
exec UspCTMAllocateDay1Room2410
exec UspCTMAllocateDay2Room2410
exec UspCTMAllocateDay3Room2410
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

create procedure test as
begin
select * from tbSTESubjectProfile
select * from tbSTERoomProfile
end 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

