if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vwSTEExaminee]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vwSTEExaminee]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vwSTEExaminee
AS
SELECT          dbo.tbSTEExamineeProfile.iExamineeProfileId, 
                      dbo.tbSTEExamineeProfile.iJukenNumber, dbo.tbSTEExamineeProfile.iNendo, 
                      dbo.tbSTEExamineeProfile.vExamineeName, dbo.tbSTEExamineeProfile.vKanaName, 
                      dbo.tbSTEZipCodeMaster.vZipCodeName, 
                      dbo.tbSTEZipCodeMaster.vPrefectureName, dbo.tbSTEZipCodeMaster.vCityName, 
                      dbo.tbSTEZipCodeMaster.vAddress1, dbo.tbSTEZipCodeMaster.vAddress2, 
                      dbo.tbSTEExamineeProfile.vAddress, dbo.tbSTEExamineeProfile.iSex, 
                      CASE dbo.tbSTEExamineeProfile.iSex WHEN 0 THEN 'M' ELSE 'F' END vSex, 
                      dbo.tbSTEHighSchoolType.vHighSchoolCode, 
                      SUBSTRING(dbo.tbSTEHighSchoolType.vHighSchoolCode, 1, 2) vHighSchoolPlace,
                       SUBSTRING(dbo.tbSTEHighSchoolType.vHighSchoolCode, 3, 1) 
                      vHighSchoolType, 
                      dbo.tbSTEHighSchoolType.vHighSchoolName AS vHighSchoolName, 
                      dbo.tbSTEHighSchoolType.iHighSchoolRecommendation, 
                      dbo.tbSTEHighSchoolType.iZipCodeId, 
                      tbSTEZipCodeMaster_1.vZipCodeName AS vHZipCodeName, 
                      tbSTEZipCodeMaster_1.vPrefectureName AS vHPrefectureName, 
                      tbSTEZipCodeMaster_1.vCityName AS vHCityName, 
                      tbSTEZipCodeMaster_1.vAddress1 AS vHAddress1, 
                      tbSTEZipCodeMaster_1.vAddress2 AS vHAddress2, 
                      dbo.tbSTEExamineeProfile.vTelephone, dbo.tbSTEExamineeProfile.vEmailAddress, 
                      dbo.tbSTEExamineeProfile.dtBirthDay, dbo.tbSTEExamineeProfile.iAbsentFlag, 
                      dbo.tbSTEExamineeProfile.iRejectFlag, dbo.tbSTEExamineeProfile.iExamineeStatus, 
                      dbo.tbSTEExamineeProfile.dtSecondExamDay, 
                      dbo.tbSTEExamineeProfile.vNationality, 
                      dbo.tbSTEExamineeProfile.iPreferenceDay1Flag, 
                      dbo.tbSTEExamineeProfile.iPreferenceDay2Flag, 
                      dbo.tbSTEExamineeProfile.iPreferenceDay3Flag, 
                      dbo.tbSTEExamineeProfile.iMultipleApplyFlag, 
                      dbo.tbSTEExamineeProfile.iAdmissionType1, 
                      dbo.tbSTEExamineeProfile.iAdmissionType2, 
                      tbSTESubjectProfile_1.vSubjectName AS vRika1SubjectName, 
                      tbSTESubjectProfile_2.vSubjectName AS vRika2SubjectName, 
                      tbSTESubjectProfile_3.vSubjectName AS vLangSubjectName, 
                      dbo.tbSTEExamineeProfile.iBackgroundId, 
                      dbo.tbSTEExamineeProfile.iUniversityType, dbo.tbSTEExamineeProfile.iFamilyId, 
                      dbo.tbSTEExamineeProfile.iParentJobCategory, 
                      dbo.tbSTEExamineeProfile.iQualificationId, 
                      tbSTELookUpTable_7.vName AS vUnivType, 
                      tbSTELookUpTable_7.iValue AS iUnivType_VAL, 
                      tbSTELookUpTable_1.vName AS vBackground, 
                      tbSTELookUpTable_1.iValue AS iBackground_VAL, 
                      tbSTELookUpTable_2.vName AS vFamily, 
                      tbSTELookUpTable_2.iValue AS iFamilyID_VAL, 
                      tbSTELookUpTable_3.vName AS vParentJob, 
                      tbSTELookUpTable_3.iValue AS iParentJob_VAL, 
                      tbSTELookUpTable_4.vName AS vSikaku, 
                      tbSTELookUpTable_4.iValue AS iSikaku_VAL, 
                      tbSTELookUpTable_6.vName AS vPhysical, 
                      tbSTELookUpTable_6.iValue AS iPhysical_VAL, 
                      dbo.tbSTEExamineeProfile.iPhysicalConditionId, 
                      tbSTELookUpTable_5.vName AS vSuisen, CASE WHEN RIGHT(CONVERT(varchar, 
                      getdate(), 12), 4) >= RIGHT(CONVERT(varchar, 
                      dbo.tbSTEExamineeProfile.dtBirthDay, 12), 4) THEN DATEDIFF(year, 
                      dbo.tbSTEExamineeProfile.dtBirthDay, GETDATE()) ELSE DATEDIFF(year, 
                      dbo.tbSTEExamineeProfile.dtBirthDay, GETDATE()) - 1 END AS iAge, 
                      dbo.tbSTERoomProfile.vRoomName, dbo.tbSTERoomProfile.iRandom, 
                      dbo.tbSTESystemProfile.iNendo AS iSystemNendo, 
                      dbo.tbSTESystemProfile.iActiveFlag, dbo.tbSTEExamineeProfile.iCourse, 
                      dbo.tbSTEExamineeProfile.iDepartment, dbo.tbSTEExamineeProfile.vHyoteiGrade, 
                      dbo.tbSTEExamineeProfile.vUnivName, dbo.tbSTEExamineeProfile.iSuisenFlagId, 
                      dbo.tbSTEExamineeProfile.iLanguageSubjProfileId, 
                      dbo.tbSTEExamineeProfile.iScienceSubjProfileId1, 
                      dbo.tbSTEExamineeProfile.iScienceSubjProfileId2, 
                      sp_L1.iDispSubID AS iLanguageSubId, sp_S1.iDispSubID AS iScienceSubId1, 
                      sp_S2.iDispSubID AS iScienceSubId2
FROM            dbo.tbSTEExamineeProfile LEFT OUTER JOIN
                      dbo.tbSTERoomProfile ON 
                      dbo.tbSTEExamineeProfile.iRoomProfileId = dbo.tbSTERoomProfile.iRoomProfileId LEFT
                       OUTER JOIN
                      dbo.tbSTESubjectProfile tbSTESubjectProfile_3 ON 
                      dbo.tbSTEExamineeProfile.iLanguageSubjProfileId = tbSTESubjectProfile_3.iSubjectProfileId
                       LEFT OUTER JOIN
                      dbo.tbSTESubjectProfile tbSTESubjectProfile_1 ON 
                      dbo.tbSTEExamineeProfile.iScienceSubjProfileId2 = tbSTESubjectProfile_1.iSubjectProfileId
                       LEFT OUTER JOIN
                      dbo.tbSTEZipCodeMaster ON 
                      dbo.tbSTEExamineeProfile.iZipCodeId = dbo.tbSTEZipCodeMaster.iZipCodeId LEFT OUTER
                       JOIN
                      dbo.tbSTESubjectProfile tbSTESubjectProfile_2 ON 
                      dbo.tbSTEExamineeProfile.iScienceSubjProfileId1 = tbSTESubjectProfile_2.iSubjectProfileId
                       LEFT OUTER JOIN
                      dbo.tbSTEHighSchoolType ON 
                      dbo.tbSTEExamineeProfile.iHighSchoolId = dbo.tbSTEHighSchoolType.iHighSchoolId
                       LEFT OUTER JOIN
                      dbo.tbSTELookUpTable tbSTELookUpTable_5 ON 
                      dbo.tbSTEExamineeProfile.iSuisenFlagId = tbSTELookUpTable_5.iLookUpTableID LEFT
                       OUTER JOIN
                      dbo.tbSTELookUpTable tbSTELookUpTable_4 ON 
                      dbo.tbSTEExamineeProfile.iQualificationId = tbSTELookUpTable_4.iLookUpTableID LEFT
                       OUTER JOIN
                      dbo.tbSTELookUpTable tbSTELookUpTable_6 ON 
                      dbo.tbSTEExamineeProfile.iPhysicalConditionId = tbSTELookUpTable_6.iLookUpTableID
                       LEFT OUTER JOIN
                      dbo.tbSTELookUpTable tbSTELookUpTable_3 ON 
                      dbo.tbSTEExamineeProfile.iParentJobCategory = tbSTELookUpTable_3.iLookUpTableID
                       LEFT OUTER JOIN
                      dbo.tbSTELookUpTable tbSTELookUpTable_2 ON 
                      dbo.tbSTEExamineeProfile.iFamilyId = tbSTELookUpTable_2.iLookUpTableID LEFT OUTER
                       JOIN
                      dbo.tbSTELookUpTable tbSTELookUpTable_1 ON 
                      dbo.tbSTEExamineeProfile.iBackgroundId = tbSTELookUpTable_1.iLookUpTableID LEFT
                       OUTER JOIN
                      dbo.tbSTELookUpTable tbSTELookUpTable_7 ON 
                      dbo.tbSTEExamineeProfile.iUniversityType = tbSTELookUpTable_7.iLookUpTableID LEFT
                       OUTER JOIN
                      dbo.tbSTEZipCodeMaster tbSTEZipCodeMaster_1 ON 
                      dbo.tbSTEHighSchoolType.iZipCodeId = tbSTEZipCodeMaster_1.iZipCodeId CROSS
                       JOIN
                      dbo.tbSTESystemProfile LEFT OUTER JOIN
                      tbSTESubjectProfile AS sp_L1 ON 
                      sp_L1.iSubjectProfileID = tbSTEExamineeProfile.iLanguageSubjProfileId LEFT OUTER
                       JOIN
                      tbSTESubjectProfile AS sp_S1 ON 
                      sp_S1.iSubjectProfileID = tbSTEExamineeProfile.iScienceSubjProfileId1 LEFT OUTER
                       JOIN
                      tbSTESubjectProfile AS sp_S2 ON 
                      sp_S2.iSubjectProfileID = tbSTEExamineeProfile.iScienceSubjProfileId2

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

