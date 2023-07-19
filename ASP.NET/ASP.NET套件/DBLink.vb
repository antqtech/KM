Imports System.Data
Imports System.Data.SqlClient
Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Net.Mail
Namespace PISCP.DB
    Public Class DBLink
        '加密程式
        Dim AES As New cAES
        '''新增帳戶資料
        Public Function usp_Add_AAcc(ByVal strA01 As String, ByVal strA02 As String, ByVal strA03 As String, ByVal strA04 As String, ByVal strA05 As String, ByVal strA06 As String, ByVal strA07 As String, ByVal strA08 As String, ByVal strA09 As String, ByVal strA10 As String, ByVal strA11 As String, ByVal strA12 As String, ByVal strA13 As String, ByVal strA14 As String, ByVal strA15 As String, ByVal strA16 As String, ByVal strA17 As String) As DataSet
            Dim strCommand As String = "usp_Add_AAcc"
            Dim strTableName As String = "AAcc"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@A02", strA02)
            dbCheck.htParam.Add("@A03", strA03)
            dbCheck.htParam.Add("@A04", strA04)
            dbCheck.htParam.Add("@A05", strA05)
            dbCheck.htParam.Add("@A06", strA06)
            dbCheck.htParam.Add("@A07", strA07)
            dbCheck.htParam.Add("@A08", strA08)
            dbCheck.htParam.Add("@A09", strA09)
            dbCheck.htParam.Add("@A10", strA10)
            dbCheck.htParam.Add("@A11", strA11)
            dbCheck.htParam.Add("@A12", strA12)
            dbCheck.htParam.Add("@A13", strA13)
            dbCheck.htParam.Add("@A14", strA14)
            dbCheck.htParam.Add("@A15", strA15)
            dbCheck.htParam.Add("@A16", strA16)
            dbCheck.htParam.Add("@A17", strA17)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function



        '''新增帳戶資料
        Public Function usp_Add_AAcc2(ByVal strA01 As String, ByVal strA02 As String, ByVal strA03 As String, ByVal strA04 As String, ByVal strA05 As String, ByVal strA06 As String, ByVal strA07 As String, ByVal strA08 As String, ByVal strA09 As String, ByVal strA10 As String, ByVal strA11 As String, ByVal strA12 As String, ByVal strA13 As String, ByVal strA14 As String, ByVal strA15 As String, ByVal strA16 As String, ByVal strA17 As String, ByVal strA18 As String, ByVal strA19 As String, ByVal strA20 As String) As DataSet
            Dim strCommand As String = "usp_Add_AAcc2"
            Dim strTableName As String = "AAcc2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@A02", strA02)
            dbCheck.htParam.Add("@A03", strA03)
            dbCheck.htParam.Add("@A04", strA04)
            dbCheck.htParam.Add("@A05", strA05)
            dbCheck.htParam.Add("@A06", strA06)
            dbCheck.htParam.Add("@A07", strA07)
            dbCheck.htParam.Add("@A08", strA08)
            dbCheck.htParam.Add("@A09", strA09)
            dbCheck.htParam.Add("@A10", strA10)
            dbCheck.htParam.Add("@A11", strA11)
            dbCheck.htParam.Add("@A12", strA12)
            dbCheck.htParam.Add("@A13", strA13)
            dbCheck.htParam.Add("@A14", strA14)
            dbCheck.htParam.Add("@A15", strA15)
            dbCheck.htParam.Add("@A16", strA16)
            dbCheck.htParam.Add("@A17", strA17)
            dbCheck.htParam.Add("@A18", strA18)
            dbCheck.htParam.Add("@A19", strA19)
            dbCheck.htParam.Add("@A20", strA20)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增教保輔導紀錄資料表
        Public Function usp_Add_AEcr1(ByVal strCRID As String, ByVal strCreateDT As String, ByVal strEntryDate As String, ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String, ByVal strTVersion As String) As DataSet
            Dim strCommand As String = "usp_Add_AEcr1"
            Dim strTableName As String = "AEcr1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@CreateDT", strCreateDT)
            dbCheck.htParam.Add("@EntryDate", strEntryDate)
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            dbCheck.htParam.Add("@TVersion", strTVersion)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增教保輔導人員資料表
        Public Function usp_Add_AEcm1(ByVal strCRID As String, ByVal strA01 As String, ByVal strGCIdentity As String) As DataSet
            Dim strCommand As String = "usp_Add_AEcm1"
            Dim strTableName As String = "AEcm1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@GCIdentity", strGCIdentity)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增教保輔導健康人數資料表
        Public Function usp_Add_AEch1(ByVal strCRID As String, ByVal strSymptom As String, ByVal strHealthNum As String) As DataSet
            Dim strCommand As String = "usp_Add_AEch1"
            Dim strTableName As String = "AEch1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@Symptom", strSymptom)
            dbCheck.htParam.Add("@HealthNum", strHealthNum)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增教保輔導優點與改善策略建議資料表(輔導紀錄編號、建議編號、內容、類型、指標版本、指標編號、追蹤狀況版本、追蹤狀況編號、最新追蹤狀況編號、改善原因版本、改善原因編號、身份、結案日期)
        Public Function usp_Add_AEcs1(ByVal strCRID As String, ByVal strCRSID As String, ByVal strSuggest As String, ByVal strSType As String,
                                      ByVal strTVersion As String, ByVal strTID As String, ByVal strTSVersion As String, ByVal strTSID As String,
                                      ByVal strNewTSID As String, ByVal strIRVersion As String, ByVal strIRID As String,
                                      ByVal strGCIdentity As String, ByVal strCloseDate As String) As DataSet
            Dim strCommand As String = "usp_Add_AEcs1"
            Dim strTableName As String = "AEcs1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@CRSID", strCRSID)
            dbCheck.htParam.Add("@Suggest", strSuggest)
            dbCheck.htParam.Add("@SType", strSType)
            dbCheck.htParam.Add("@TVersion", strTVersion)
            dbCheck.htParam.Add("@TID", strTID)
            dbCheck.htParam.Add("@TSVersion", strTSVersion)
            dbCheck.htParam.Add("@TSID", strTSID)
            dbCheck.htParam.Add("@NewTSID", strNewTSID)
            dbCheck.htParam.Add("@IRVersion", strIRVersion)
            dbCheck.htParam.Add("@IRID", strIRID)
            dbCheck.htParam.Add("@GCIdentity", strGCIdentity)
            dbCheck.htParam.Add("@CloseDate", strCloseDate)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增教保輔導前次建議追蹤資料表
        Public Function usp_Add_AOld1(ByVal strCRID As String, ByVal strOldCRID As String, ByVal strOldCRSID As String, ByVal strCloseDate As String) As DataSet
            Dim strCommand As String = "usp_Add_AOld1"
            Dim strTableName As String = "AOld1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@OldCRID", strOldCRID)
            dbCheck.htParam.Add("@OldCRSID", strOldCRSID)
            dbCheck.htParam.Add("@CloseDate", strCloseDate)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增園所資料表(縣市編號、園所編號、園所名稱、是否有效、公私立、園所地址、園所電話、備註)
        Public Function usp_Add_AGtn1(ByVal strCityID As String, ByVal strGID As String, ByVal strGName As String, ByVal strGEff As String, ByVal strGPP As String, ByVal strGAddress As String, ByVal strGTEL As String, ByVal strGRemark As String) As DataSet
            Dim strCommand As String = "usp_Add_AGtn1"
            Dim strTableName As String = "AGtn1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GName", strGName)
            dbCheck.htParam.Add("@GEff", strGEff)
            dbCheck.htParam.Add("@GPP", strGPP)
            dbCheck.htParam.Add("@GAddress", strGAddress)
            dbCheck.htParam.Add("@GTEL", strGTEL)
            dbCheck.htParam.Add("@GRemark", strGRemark)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增指標版本、指標編號(層面)、指標名稱、指標類型、是否使用此版本
        Public Function usp_Add_ATgt1(ByVal strTVersion As String, ByVal strTID As String, ByVal strTName As String, ByVal strTType As String, ByVal strTYorN As String) As DataSet
            Dim strCommand As String = "usp_Add_ATgt1"
            Dim strTableName As String = "ATgt1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TVersion", strTVersion)
            dbCheck.htParam.Add("@TID", strTID)
            dbCheck.htParam.Add("@TName", strTName)
            dbCheck.htParam.Add("@TType", strTType)
            dbCheck.htParam.Add("@TYorN", strTYorN)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增追蹤狀況版本、追蹤狀況編號、追蹤狀況名稱、追蹤狀況定義、是否使用此版本
        Public Function usp_Add_ATst1(ByVal strTSVersion As String, ByVal strTSID As String, ByVal strTSName As String, ByVal strTSDef As String, ByVal strTSYorN As String) As DataSet
            Dim strCommand As String = "usp_Add_ATst1"
            Dim strTableName As String = "ATst1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TSVersion", strTSVersion)
            dbCheck.htParam.Add("@TSID", strTSID)
            dbCheck.htParam.Add("@TSName", strTSName)
            dbCheck.htParam.Add("@TSDef", strTSDef)
            dbCheck.htParam.Add("@TSYorN", strTSYorN)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增改善原因版本、改善原因編號、改善原因名稱、改善原因定義、、是否使用此版本1 是 0 否
        Public Function usp_Add_AIpr1(ByVal strIRVersion As String, ByVal strIRID As String, ByVal strIRName As String, ByVal strIRDef As String, ByVal strIRYorN As String) As DataSet
            Dim strCommand As String = "usp_Add_AIpr1"
            Dim strTableName As String = "AIpr1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@IRVersion", strIRVersion)
            dbCheck.htParam.Add("@IRID", strIRID)
            dbCheck.htParam.Add("@IRName", strIRName)
            dbCheck.htParam.Add("@IRDef", strIRDef)
            dbCheck.htParam.Add("@IRYorN", strIRYorN)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增追蹤狀況版本、追蹤狀況編號、追蹤狀況名稱、追蹤狀況定義
        Public Function usp_Add_AGcm1(ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String, ByVal strA01 As String, ByVal strGCIdentity As String) As DataSet
            Dim strCommand As String = "usp_Add_AGcm1"
            Dim strTableName As String = "AGcm1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@GCIdentity", strGCIdentity)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增園所班級資料表(縣市編號、園所編號、班級編號、班級名稱、是否有效、班別、2~3歲人數、3~4歲人數、4~5歲人數、5歲人數)
        Public Function usp_Add_AGtc1(ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String, ByVal strGCName As String, ByVal strGCEff As String, ByVal strGCType As String, ByVal strGCPeople2 As String, ByVal strGCPeople3 As String, ByVal strGCPeople4 As String, ByVal strGCPeople5 As String) As DataSet
            Dim strCommand As String = "usp_Add_AGtc1"
            Dim strTableName As String = "AGtc1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            dbCheck.htParam.Add("@GCName", strGCName)
            dbCheck.htParam.Add("@GCEff", strGCEff)
            dbCheck.htParam.Add("@GCType", strGCType)
            dbCheck.htParam.Add("@GCPeople2", strGCPeople2)
            dbCheck.htParam.Add("@GCPeople3", strGCPeople3)
            dbCheck.htParam.Add("@GCPeople4", strGCPeople4)
            dbCheck.htParam.Add("@GCPeople5", strGCPeople5)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增錯誤紀錄資料表
        Public Function usp_Add_AErr1(ByVal strErrTime As String, ByVal strErrPage As String, ByVal strErrCon As String, ByVal strErrMsg As String) As DataSet
            Dim strCommand As String = "usp_Add_AErr1"
            Dim strTableName As String = "AErr1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@ErrTime", strErrTime)
            dbCheck.htParam.Add("@ErrPage", strErrPage)
            dbCheck.htParam.Add("@ErrCon", strErrCon)
            dbCheck.htParam.Add("@ErrMsg", strErrMsg)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增日期
        Public Function usp_Add_ADateDateDate(ByVal strDate1 As String, ByVal strKeyy As String, ByVal strVII As String) As DataSet
            Dim strCommand As String = "usp_Add_ADateDateDate"
            Dim strTableName As String = "ADateDateDate"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@Date1", strDate1)
            dbCheck.htParam.Add("@Keyy", strKeyy)
            dbCheck.htParam.Add("@VII", strVII)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增系統公告資料表
        Public Function usp_Add_ANtF1(ByVal strTimes As String, ByVal strTopic As String, ByVal strContents As String) As DataSet
            Dim strCommand As String = "usp_Add_ANtF1"
            Dim strTableName As String = "ANtF1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@Times", strTimes)
            dbCheck.htParam.Add("@Topic", strTopic)
            dbCheck.htParam.Add("@Contents", strContents)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增新增問卷管理資料表(問卷編號、問卷名稱、問卷說明、開放填答開始時間、開放填答最後時間)
        Public Function usp_Add_AQme1(ByVal strQID As String, ByVal strQName As String, ByVal strQContents As String, ByVal strQStarttime As String, ByVal strQEndtime As String) As DataSet
            Dim strCommand As String = "usp_Add_AQme1"
            Dim strTableName As String = "AQme1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QName", strQName)
            dbCheck.htParam.Add("@QContents", strQContents)
            dbCheck.htParam.Add("@QStarttime", strQStarttime)
            dbCheck.htParam.Add("@QEndtime", strQEndtime)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增問卷題目分類資料表(問卷編號、題目分類編號、題目分類名稱、題目分類說明)
        Public Function usp_Add_AQcf1(ByVal strQID As String, ByVal strQCID As String, ByVal strQCName As String, ByVal strQCcontents As String) As DataSet
            Dim strCommand As String = "usp_Add_AQcf1"
            Dim strTableName As String = "AQcf1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            dbCheck.htParam.Add("@QCName", strQCName)
            dbCheck.htParam.Add("@QCcontents", strQCcontents)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增問卷題目(問卷編號、題目分類編號、題目編號、題目名稱、選項類型、必填、跳題、其他)
        Public Function usp_Add_AQue1(ByVal strQID As String, ByVal strQCID As String, ByVal strQTID As String,
                                      ByVal strQTName As String, ByVal strQTtype As String, ByVal strQTchoice As String,
                                      ByVal strQTOther As String) As DataSet
            Dim strCommand As String = "usp_Add_AQue1"
            Dim strTableName As String = "AQue1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            dbCheck.htParam.Add("@QTID", strQTID)
            dbCheck.htParam.Add("@QTName", strQTName)
            dbCheck.htParam.Add("@QTtype", strQTtype)
            dbCheck.htParam.Add("@QTchoice", strQTchoice)
            dbCheck.htParam.Add("@QTOther", strQTOther)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''新增問卷選項(問卷編號、題目分類編號、題目編號、選項編號、選項名稱、跳題)
        Public Function usp_Add_AQop1(ByVal strQID As String, ByVal strQCID As String, ByVal strQTID As String, ByVal strQOID As String, ByVal strQOName As String, ByVal strQOjump As String) As DataSet
            Dim strCommand As String = "usp_Add_AQop1"
            Dim strTableName As String = "AQop1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            dbCheck.htParam.Add("@QTID", strQTID)
            dbCheck.htParam.Add("@QOID", strQOID)
            dbCheck.htParam.Add("@QOName", strQOName)
            dbCheck.htParam.Add("@QOjump", strQOjump)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增問卷填寫紀錄(問卷編號、題目分類編號、題目編號、選項編號、帳號、其他文字)
        Public Function usp_Add_AQrc1(ByVal strQID As String, ByVal strQCID As String, ByVal strQTID As String,
                                      ByVal strQOID As String, ByVal strA01 As String, ByVal strQAns As String) As DataSet
            Dim strCommand As String = "usp_Add_AQrc1"
            Dim strTableName As String = "AQrc1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            dbCheck.htParam.Add("@QTID", strQTID)
            dbCheck.htParam.Add("@QOID", strQOID)
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@QAns", strQAns)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增問卷受訪人員(問卷編號、帳號、填答時間、狀態)
        Public Function usp_Add_AQur1(ByVal strQID As String, ByVal strA01 As String, ByVal strQTime As String,
                                      ByVal strQState As String, ByVal strIPAddress As String) As DataSet
            Dim strCommand As String = "usp_Add_AQur1"
            Dim strTableName As String = "AQur1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@QTime", strQTime)
            dbCheck.htParam.Add("@QState", strQState)
            dbCheck.htParam.Add("@IPAddress", strIPAddress)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增月輔導報告資料表(縣市編號、報告日期、應上班天數、應輔導天數、教保輔導天數、非教保輔導天數、需要協助或建議、縣市政府之相關意見及處理策略、巡輔員帳號)
        Public Function usp_Add_ARecrm1(ByVal strCityID As String, ByVal strReportDate As String, ByVal strWorkDay As String,
                                      ByVal strCoachDay As String, ByVal strEduCRDay As String, ByVal strNotEduCRDay As String,
                                         ByVal strAssistance As String, ByVal strGovAssistance As String, ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Add_ARecrm1"
            Dim strTableName As String = "ARecrm1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@WorkDay", strWorkDay)
            dbCheck.htParam.Add("@CoachDay", strCoachDay)
            dbCheck.htParam.Add("@EduCRDay", strEduCRDay)
            dbCheck.htParam.Add("@NotEduCRDay", strNotEduCRDay)
            dbCheck.htParam.Add("@Assistance", strAssistance)
            dbCheck.htParam.Add("@GovAssistance", strGovAssistance)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增每月非教保輔導資料表(縣市編號、報告日期、項目編號、日期、開始時間、結束時間、機構名稱、處理業務簡述、備註、巡輔員帳號)
        Public Function usp_Add_ARnecrm1(ByVal strCityID As String, ByVal strReportDate As String, ByVal strReportID As String,
                                      ByVal strDay As String, ByVal strStartTime As String, ByVal strEndTime As String,
                                         ByVal strInstitution As String, ByVal strBusiness As String, ByVal strReportRemark As String,
                                         ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Add_ARnecrm1"
            Dim strTableName As String = "ARnecrm1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@ReportID", strReportID)
            dbCheck.htParam.Add("@Day", strDay)
            dbCheck.htParam.Add("@StartTime", strStartTime)
            dbCheck.htParam.Add("@EndTime", strEndTime)
            dbCheck.htParam.Add("@Institution", strInstitution)
            dbCheck.htParam.Add("@Business", strBusiness)
            dbCheck.htParam.Add("@ReportRemark", strReportRemark)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增輔導教授每月填報意見資料表(縣市編號、報告日期、需要協助或建議、縣市政府之相關意見及處理策略、巡輔員提供列印功能、設定列印的巡輔員帳號、輔導教授帳號)
        Public Function usp_Add_AReportPA1(ByVal strCityID As String, ByVal strReportDate As String, ByVal strAssistance As String,
                                           ByVal strGovAssistance As String, ByVal strPrintAut As String, ByVal strPrintA01 As String,
                                            ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Add_AReportPA1"
            Dim strTableName As String = "AReportPA1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@Assistance", strAssistance)
            dbCheck.htParam.Add("@GovAssistance", strGovAssistance)
            dbCheck.htParam.Add("@PrintAut", strPrintAut)
            dbCheck.htParam.Add("@PrintA01", strPrintA01)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增輔導教授每月填報園所資料表(縣市編號、報告日期、項目編號、日期、園所編號、內容(訪視心得/省思)、輔導教授帳號)
        Public Function usp_Add_AReportPG1(ByVal strCityID As String, ByVal strReportDate As String, ByVal strReportID As String,
                                      ByVal strDay As String, ByVal strGID As String, ByVal strPContent As String, ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Add_AReportPG1"
            Dim strTableName As String = "AReportPG1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@ReportID", strReportID)
            dbCheck.htParam.Add("@Day", strDay)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@PContent", strPContent)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	新增每月上傳檔案資料(縣市編號、報告日期、檔案編號、類型、檔案名稱、副檔名、巡輔員帳號)
        Public Function usp_Add_AUlf1(ByVal strCityID As String, ByVal strReportDate As String, ByVal strFileID As String,
                                      ByVal strType As String, ByVal strFileName As String, ByVal strFileName2 As String,
                                      ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Add_AUlf1"
            Dim strTableName As String = "AUlf1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@FileID", strFileID)
            dbCheck.htParam.Add("@Type", strType)
            dbCheck.htParam.Add("@FileName", strFileName)
            dbCheck.htParam.Add("@FileName2", strFileName2)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	新增每月會議記錄出席教授資料(縣市編號、報告日期、編號、輔導教授帳號、日期、會議名稱、巡輔員帳號)
        Public Function usp_Add_AAmt1(ByVal strCityID As String, ByVal strReportDate As String, ByVal strAMID As String,
                                      ByVal strA01 As String, ByVal strDay As String, ByVal strMeetName As String,
                                      ByVal strAA01 As String) As DataSet
            Dim strCommand As String = "usp_Add_AAmt1"
            Dim strTableName As String = "AAmt1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@AMID", strAMID)
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@Day", strDay)
            dbCheck.htParam.Add("@MeetName", strMeetName)
            dbCheck.htParam.Add("@AA01", strAA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''		新增行政輔導紀錄資料表((回報日期、前次輔導紀錄編號、前次建議編號、縣市、幼兒園、班級、追蹤狀況編號、改善原因編號、追蹤方式、內容))
        Public Function usp_Add_AAmr1(ByVal strReportDate As String, ByVal strOldCRID As String, ByVal strOldCRSID As String,
                                      ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String, ByVal strTSID As String,
                                       ByVal strIRID As String, ByVal strTrackWay As String, ByVal strContents As String) As DataSet
            Dim strCommand As String = "usp_Add_AAmr1"
            Dim strTableName As String = "AAmr1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@OldCRID", strOldCRID)
            dbCheck.htParam.Add("@OldCRSID", strOldCRSID)
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            dbCheck.htParam.Add("@TSID", strTSID)
            dbCheck.htParam.Add("@IRID", strIRID)
            dbCheck.htParam.Add("@TrackWay", strTrackWay)
            dbCheck.htParam.Add("@Contents", strContents)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢帳號、姓名、生日、性別、手機號碼、聯絡電話、電子信箱、教師身份、園所代碼、註冊日期、使用者身份、Key、VI、權限
        Public Function usp_Sel_SAcc1(ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc1"
            Dim strTableName As String = "SAcc1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''檢查密碼。查詢密碼、Key、VI、權限、密碼錯誤次數、帳戶停用訊息代號
        Public Function usp_Sel_SAccK(ByVal strA01 As String, ByVal strA02 As String) As String
            Dim strCommand1 As String = "usp_Sel_SAccK"
            Dim strTableName1 As String = "SAccK"
            Dim dbCheck1 As New PISCP.DB.ClassDB
            Dim dsDone1 As New DataSet()
            Dim spParams1 As SqlParameter() = Nothing

            dbCheck1.htParam.Clear()
            dbCheck1.htParam.Add("@A01", strA01)
            spParams1 = dbCheck1.getSpParams(0)
            dsDone1 = dbCheck1.RunProcedure(strCommand1, spParams1, strTableName1)

            If dsDone1.Tables(0).Rows.Count <> 0 Then
                '解密(密碼)
                Dim pass As String
                pass = AES.AES_Decrypt(dsDone1.Tables(0).Rows(0).Item(0).ToString(), dsDone1.Tables(0).Rows(0).Item(1).ToString(), dsDone1.Tables(0).Rows(0).Item(2).ToString())

                '修改密碼錯誤次數
                Dim strCommand2 As String = "usp_Upd_UAcc2"
                Dim strTableName2 As String = "UAcc2"
                Dim dbCheck2 As New PISCP.DB.ClassDB
                Dim spParams2 As SqlParameter() = Nothing

                '查詢停用訊息
                Dim strCommand4 As String = "usp_Sel_SAcc2"
                Dim strTableName4 As String = "SAcc2"
                Dim dbCheck4 As New PISCP.DB.ClassDB
                Dim dsDone4 As New DataSet()
                Dim spParams4 As SqlParameter() = Nothing

                '判斷密碼正確性
                If strA02 = pass Then
                    If dsDone1.Tables(0).Rows(0).Item(3).ToString() = "0" Then
                        '查詢停用訊息
                        dbCheck4.htParam.Clear()
                        dbCheck4.htParam.Add("@DisMsgID", Int(dsDone1.Tables(0).Rows(0).Item(5).ToString()))
                        spParams4 = dbCheck4.getSpParams(0)
                        dsDone4 = dbCheck4.RunProcedure(strCommand4, spParams4, strTableName4)

                        Return "Fail3&&&" & dsDone4.Tables(0).Rows(0).Item(0).ToString() & "&&&"
                    Else
                        '修改密碼錯誤次數(歸0)
                        dbCheck2.htParam.Clear()
                        dbCheck2.htParam.Add("@A15", "0")
                        dbCheck2.htParam.Add("@A01", strA01)
                        spParams2 = dbCheck2.getSpParams(0)
                        dbCheck2.RunProcedure(strCommand2, spParams2, strTableName2)

                        Return "Correct&&&密碼輸入正確~&&&"
                    End If
                Else
                    '判斷密碼輸入錯誤次數
                    If Int(dsDone1.Tables(0).Rows(0).Item(4).ToString()) + 1 >= 5 Then
                        If Int(dsDone1.Tables(0).Rows(0).Item(4).ToString()) <> 5 Then
                            '修改密碼錯誤次數(5次)
                            dbCheck2.htParam.Clear()
                            dbCheck2.htParam.Add("@A15", "5")
                            dbCheck2.htParam.Add("@A01", strA01)
                            spParams2 = dbCheck2.getSpParams(0)
                            dbCheck2.RunProcedure(strCommand2, spParams2, strTableName2)
                        End If

                        If dsDone1.Tables(0).Rows(0).Item(3).ToString() <> "0" Then
                            '停用帳戶
                            Dim strCommand3 As String = "usp_Upd_UAcc3"
                            Dim strTableName3 As String = "UAcc3"
                            Dim dbCheck3 As New PISCP.DB.ClassDB
                            Dim spParams3 As SqlParameter() = Nothing

                            dbCheck3.htParam.Clear()
                            dbCheck3.htParam.Add("@A19", "0")
                            dbCheck3.htParam.Add("@A16", "2")
                            dbCheck3.htParam.Add("@A01", strA01)
                            spParams3 = dbCheck3.getSpParams(0)
                            dbCheck3.RunProcedure(strCommand3, spParams3, strTableName3)
                        End If
                        '查詢停用訊息
                        dbCheck4.htParam.Clear()
                        dbCheck4.htParam.Add("@DisMsgID", "2")
                        spParams4 = dbCheck4.getSpParams(0)
                        dsDone4 = dbCheck4.RunProcedure(strCommand4, spParams4, strTableName4)

                        Return "Fail3&&&" & dsDone4.Tables(0).Rows(0).Item(0).ToString() & "&&&"
                    Else
                        '修改密碼錯誤次數
                        dbCheck2.htParam.Clear()
                        dbCheck2.htParam.Add("@A15", Int(dsDone1.Tables(0).Rows(0).Item(4).ToString()) + 1)
                        dbCheck2.htParam.Add("@A01", strA01)
                        spParams2 = dbCheck2.getSpParams(0)
                        dbCheck2.RunProcedure(strCommand2, spParams2, strTableName2)

                        Return "Fail2&&&您的密碼連續輸入錯誤" & Int(dsDone1.Tables(0).Rows(0).Item(4).ToString()) + 1 & "次，若連續輸入錯誤達到5次，您的帳戶將停用！&&&"
                    End If
                End If
            Else
                Return "Fail1&&&查無帳號，請先註冊！&&&"
            End If
        End Function

        '''查詢帳號、姓名、Key、VI(權限C)
        Public Function usp_Sel_SAcc3(ByVal strA14 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc3"
            Dim strTableName As String = "SAcc3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A14", strA14)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢帳號、姓名、Key、VI
        Public Function usp_Sel_SAcc4(ByVal strA14 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc4"
            Dim strTableName As String = "SAcc4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A14", strA14)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢帳號、姓名、教師身份、園所代碼、使用者身份、Key、VI、權限
        Public Function usp_Sel_SAcc5() As DataSet
            Dim strCommand As String = "usp_Sel_SAcc5"
            Dim strTableName As String = "SAcc5"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''查詢帳號、姓名、生日、性別、手機號碼、聯絡電話、電子信箱、教師身份、園所代碼、註冊日期、使用者身份、備註、Key、VI、停用訊息代號、權限、巡輔員專任/兼任
        Public Function usp_Sel_SAcc6(ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc6"
            Dim strTableName As String = "SAcc6"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢帳號、姓名、幼兒園名稱、教師身份、使用者身份、權限、Key、VI、順序編號、縣市編號
        Public Function usp_Sel_SAcc7(ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc7"
            Dim strTableName As String = "SAcc7"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((帳號、縣市))帳號、姓名、幼兒園名稱、教師身份、使用者身份、權限、Key、VI、順序編號、縣市編號
        Public Function usp_Sel_SAcc8(ByVal strA01 As String, ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc8"
            Dim strTableName As String = "SAcc8"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：姓名))帳號、姓名、幼兒園名稱、教師身份、使用者身份、權限、Key、VI、順序編號、縣市編號
        Public Function usp_Sel_SAcc9(ByVal strA05 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc9"
            Dim strTableName As String = "SAcc9"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A05", strA05)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：縣市、姓名))帳號、姓名、幼兒園名稱、教師身份、使用者身份、權限、Key、VI、順序編號、縣市編號
        Public Function usp_Sel_SAcc10(ByVal strCityID As String, ByVal strA05 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc10"
            Dim strTableName As String = "SAcc10"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@A05", strA05)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：帳號、姓名))帳號、姓名、幼兒園名稱、教師身份、使用者身份、權限、Key、VI、順序編號、縣市編號
        Public Function usp_Sel_SAcc11(ByVal strA01 As String, ByVal strA05 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc11"
            Dim strTableName As String = "SAcc11"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@A05", strA05)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：帳號、縣市、姓名))帳號、姓名、幼兒園名稱、教師身份、使用者身份、權限、Key、VI、順序編號、縣市編號
        Public Function usp_Sel_SAcc12(ByVal strA01 As String, ByVal strCityID As String, ByVal strA05 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc12"
            Dim strTableName As String = "SAcc12"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@A05", strA05)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢電子信箱
        Public Function usp_Sel_SAcc13(ByVal strA10 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc13"
            Dim strTableName As String = "SAcc13"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A10", strA10)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢帳號
        Public Function usp_Sel_SAcc14(ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc14"
            Dim strTableName As String = "SAcc14"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：帳號))帳號、姓名、教師身份、園所代碼、使用者身份、Key、VI、權限
        Public Function usp_Sel_SAcc15(ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc15"
            Dim strTableName As String = "SAcc15"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：縣市))帳號、姓名、教師身份、園所代碼、使用者身份、Key、VI、權限
        Public Function usp_Sel_SAcc16(ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc16"
            Dim strTableName As String = "SAcc16"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：帳號、縣市))帳號、姓名、教師身份、園所代碼、使用者身份、Key、VI、權限
        Public Function usp_Sel_SAcc17(ByVal strCityID As String, ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc17"
            Dim strTableName As String = "SAcc17"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢帳號、姓名、生日、性別、手機號碼、聯絡電話、電子信箱、教師身份、園所代碼、註冊日期、帳戶停用訊息代號、權限、備註、Key、VI
        Public Function usp_Sel_SAcc18() As DataSet
            Dim strCommand As String = "usp_Sel_SAcc18"
            Dim strTableName As String = "SAcc18"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：GCM帳號))帳號、姓名、教師身份、園所代碼、使用者身份、Key、VI、權限
        Public Function usp_Sel_SAcc19(ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc19"
            Dim strTableName As String = "SAcc19"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：姓名))帳號、姓名、教師身份、園所代碼、使用者身份、Key、VI、權限
        Public Function usp_Sel_SAcc20(ByVal strA05 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc20"
            Dim strTableName As String = "SAcc20"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A05", strA05)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢帳號、姓名、Key、VI
        Public Function usp_Sel_SAcc21(ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc21"
            Dim strTableName As String = "SAcc21"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''((檢查身分證重複用))查詢身分證、Key、VI
        Public Function usp_Sel_SAcc22() As DataSet
            Dim strCommand As String = "usp_Sel_SAcc22"
            Dim strTableName As String = "SAcc22"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''((檢查身分證重複用))查詢身分證、Key、VI
        Public Function usp_Sel_SAcc23(ByVal strA14 As String, ByVal strA19 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc23"
            Dim strTableName As String = "SAcc23"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A14", strA14)
            dbCheck.htParam.Add("@A19", strA19)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((自由條件：：縣市、GCM帳號、帳號、權限))帳號、姓名、教師身份、園所代碼、使用者身份、Key、VI、權限
        Public Function usp_Sel_SAcc24(ByVal strCityID As String, ByVal strA01 As String, ByVal strAA01 As String, ByVal strA19 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc24"
            Dim strTableName As String = "SAcc24"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@AA01", strAA01)
            dbCheck.htParam.Add("@A19", strA19)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢ACT 資料表
        Public Function usp_Sel_SAcc25() As DataSet
            Dim strCommand As String = "usp_Sel_SAcc25"
            Dim strTableName As String = "SAcc25"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢 縣市政府人員，帳號所設定的縣市
        Public Function usp_Sel_SAcc26(ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc26"
            Dim strTableName As String = "SAcc26"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢 屬於 縣市政府人員的 縣市的巡輔員((條件：巡輔員C、縣市))
        Public Function usp_Sel_SAcc27(ByVal strA14 As String, ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAcc27"
            Dim strTableName As String = "SAcc27"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A14", strA14)
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢ACT資料表
        Public Function usp_Sel_SACCCCCCCCCC(ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SACCCCCCCCCC"
            Dim strTableName As String = "SACCCCCCCCCC"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢日期
        Public Function usp_Sel_SADDDDDDDate() As DataSet
            Dim strCommand As String = "usp_Sel_SADDDDDDDate"
            Dim strTableName As String = "SADDDDDDDate"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢停用訊息
        Public Function usp_Sel_SMsg1() As DataSet
            Dim strCommand As String = "usp_Sel_SMsg1"
            Dim strTableName As String = "SMsg1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢縣市編號、縣市名稱
        Public Function usp_Sel_SCgp1(ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SCgp1"
            Dim strTableName As String = "SCgp1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢縣市編號、縣市名稱、順序編碼
        Public Function usp_Sel_SCgp2() As DataSet
            Dim strCommand As String = "usp_Sel_SCgp2"
            Dim strTableName As String = "SCgp2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢縣市編號、縣市名稱、順序編碼
        Public Function usp_Sel_SCgp3(ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SCgp3"
            Dim strTableName As String = "SCgp3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢園所編號、園所名稱
        Public Function usp_Sel_SGtn1(ByVal strA01 As String, ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGtn1"
            Dim strTableName As String = "SGtn1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢縣市編號、園所編號、園所名稱、是否有效、公私立、園所地址、園所電話、備註
        Public Function usp_Sel_SGtn2() As DataSet
            Dim strCommand As String = "usp_Sel_SGtn2"
            Dim strTableName As String = "SGtn2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢(條件縣市)縣市編號、園所編號、園所名稱、是否有效、公私立、園所地址、園所電話、備註
        Public Function usp_Sel_SGtn3(ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGtn3"
            Dim strTableName As String = "SGtn3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢(條件縣市)縣市編號、園所編號、園所名稱、是否有效、公私立、園所地址、園所電話、備註
        Public Function usp_Sel_SGtn4(ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGtn4"
            Dim strTableName As String = "SGtn4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢(條件縣市)縣市編號、園所編號、園所名稱、是否有效、公私立、園所地址、園所電話、備註
        Public Function usp_Sel_SGtn5(ByVal strGName As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGtn5"
            Dim strTableName As String = "SGtn5"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@GName", strGName)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢(條件縣市與巡輔員)縣市編號、園所編號、園所名稱、是否有效、公私立、園所地址、園所電話、備註
        Public Function usp_Sel_SGtn6(ByVal strCityID As String, ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGtn6"
            Dim strTableName As String = "SGtn6"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢(條件縣市與幼兒園)縣市編號、園所編號、園所名稱、是否有效、公私立、園所地址、園所電話、備註
        Public Function usp_Sel_SGtn7(ByVal strCityID As String, ByVal strGName As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGtn7"
            Dim strTableName As String = "SGtn7"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GName", strGName)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢(條件巡輔員與幼兒園)縣市編號、園所編號、園所名稱、是否有效、公私立、園所地址、園所電話、備註
        Public Function usp_Sel_SGtn8(ByVal strA01 As String, ByVal strGName As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGtn8"
            Dim strTableName As String = "SGtn8"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@GName", strGName)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢(條件縣市與巡輔員與幼兒園)縣市編號、園所編號、園所名稱、是否有效、公私立、園所地址、園所電話、備註
        Public Function usp_Sel_SGtn9(ByVal strCityID As String, ByVal strA01 As String, ByVal strGName As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGtn9"
            Dim strTableName As String = "SGtn9"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@GName", strGName)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢最大園所編號
        Public Function usp_Sel_SGtn10(ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGtn10"
            Dim strTableName As String = "SGtn10"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢最大園所編號
        Public Function usp_Sel_SGtn11(ByVal strCityID As String, ByVal strGID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGtn11"
            Dim strTableName As String = "SGtn11"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID ", strGID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((自由條件：縣市、帳號、園所名稱、是否有效))縣市編號、園所編號、園所名稱、是否有效、公私立、園所地址、園所電話、備註
        Public Function usp_Sel_SGtn12(ByVal strCityID As String, ByVal strA01 As String, ByVal strGName As String, ByVal strGEff As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGtn12"
            Dim strTableName As String = "SGtn12"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@GName", strGName)
            dbCheck.htParam.Add("@GEff", strGEff)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢[即時] 幼兒園人數統計表
        Public Function usp_Sel_SGtn13() As DataSet
            Dim strCommand As String = "usp_Sel_SGtn13"
            Dim strTableName As String = "SGtn13"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢班級編號、班級名稱、帳號、園所編號
        Public Function usp_Sel_SGtc1(ByVal strA01 As String, ByVal strCityID As String, ByVal strGID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGtc1"
            Dim strTableName As String = "SGtc1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢(條件縣市與園所)縣市編號、園所編號、班級編號、班級名稱、是否有效、班別、2~3歲人數、3~4歲人數、4~5歲人數、5歲人數、帳號、身份
        Public Function usp_Sel_SGtc2(ByVal strCityID As String, ByVal strGID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGtc2"
            Dim strTableName As String = "SGtc2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢(條件縣市與園所與班級)帳號、姓名、Key、Vi
        Public Function usp_Sel_SGtc3(ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String, ByVal strGCIdentity As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGtc3"
            Dim strTableName As String = "SGtc3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            dbCheck.htParam.Add("@GCIdentity", strGCIdentity)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢(條件縣市與園所)帳號、姓名、教師身份、Key、VI
        Public Function usp_Sel_SGtc4(ByVal strA12 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGtc4"
            Dim strTableName As String = "SGtc4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A12", strA12)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢(條件縣市、園所、班級編號)縣市編號、園所編號、班級編號、班級名稱、是否有效、班別、2~3歲人數、3~4歲人數、4~5歲人數、5歲人數
        Public Function usp_Sel_SGtc5(ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGtc5"
            Dim strTableName As String = "SGtc5"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢最大班級編號
        Public Function usp_Sel_SGtc6(ByVal strCityID As String, ByVal strGID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGtc6"
            Dim strTableName As String = "SGtc6"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢姓名、Key、VI、帳號
        Public Function usp_Sel_SGcm1(ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String, ByVal strGCIdentity As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGcm1"
            Dim strTableName As String = "SGcm1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            dbCheck.htParam.Add("@GCIdentity", strGCIdentity)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢查詢帳號身份
        Public Function usp_Sel_SGcm2(ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String, ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGcm2"
            Dim strTableName As String = "SGcm2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢縣市編號、縣市名稱
        Public Function usp_Sel_SGcm3(ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGcm3"
            Dim strTableName As String = "SGcm3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：帳號))園所班級人員全部資料
        Public Function usp_Sel_SGcm4(ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGcm4"
            Dim strTableName As String = "SGcm4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：帳號(GCM權限)))縣市編號、園所編號
        Public Function usp_Sel_SGcm5(ByVal strA01 As String, ByVal strGCIdentity As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGcm5"
            Dim strTableName As String = "SGcm5"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@GCIdentity", strGCIdentity)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：縣市編號、園所編號(巡迴輔導員C)))縣市編號、園所編號
        Public Function usp_Sel_SGcm6(ByVal strCityID As String, ByVal strGID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGcm6"
            Dim strTableName As String = "SGcm6"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：帳號、GCM權限))查詢縣市編號、縣市名稱
        Public Function usp_Sel_SGcm7(ByVal strA01 As String, ByVal strGCIdentity As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGcm7"
            Dim strTableName As String = "SGcm7"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@GCIdentity", strGCIdentity)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：縣市、身分(巡輔員C、輔導教授E)))查詢縣市編號、帳號、身分
        Public Function usp_Sel_SGcm8(ByVal strCityID As String, ByVal strGCIdentity As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGcm8"
            Dim strTableName As String = "SGcm8"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GCIdentity", strGCIdentity)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：縣市))查詢縣市編號、幼兒園編號、班級編號
        Public Function usp_Sel_SGcm9(ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGcm9"
            Dim strTableName As String = "SGcm9"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢 即時教保輔導人員明細 查詢縣市、幼兒園、班級、身分別、教保服務人員、帳號、巡迴輔導員
        Public Function usp_Sel_SGcm10() As DataSet
            Dim strCommand As String = "usp_Sel_SGcm10"
            Dim strTableName As String = "SGcm10"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢((條件：縣市、園所是否有效、園所公私立))計算公立、私立園所、總園所數量
        Public Function usp_Sel_SGcm11(ByVal strCityID As String, ByVal strGEff As String, ByVal strGPP As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGcm11"
            Dim strTableName As String = "SGcm11"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GEff", strGEff)
            dbCheck.htParam.Add("@GPP", strGPP)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢(((條件：縣市、園所是否有效 1:有效 0:無效、園所公私立 1:公 2:私 0:不進判斷、班級有效無效 1:有效 0:無效 ))計算班級、計算班級、計算總班級
        Public Function usp_Sel_SGcm12(ByVal strCityID As String, ByVal strGEff As String, ByVal strGPP As String, ByVal strGCEff As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGcm12"
            Dim strTableName As String = "SGcm12"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GEff", strGEff)
            dbCheck.htParam.Add("@GPP", strGPP)
            dbCheck.htParam.Add("@GCEff", strGCEff)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢((條件：縣市、園所是否有效[1:有效 0:無效]、園所公私立 [1:公 2:私、0不進判斷]、班級有效無效[1:有效 0:無效]、身分[1教師、3教保員、4助理教保員] ))計算公立/私立教師、公立/私立教保員、公立/私立助理教保員
        Public Function usp_Sel_SGcm13(ByVal strCityID As String, ByVal strGEff As String, ByVal strGPP As String, ByVal strGCEff As String, ByVal strGCIdentity As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGcm13"
            Dim strTableName As String = "SGcm13"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GEff", strGEff)
            dbCheck.htParam.Add("@GPP", strGPP)
            dbCheck.htParam.Add("@GCEff", strGCEff)
            dbCheck.htParam.Add("@GCIdentity", strGCIdentity)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢 園所、班級明細表 查詢縣市、公立/私立、幼兒園名稱、班級名稱、巡迴輔導員
        Public Function usp_Sel_SGcm14() As DataSet
            Dim strCommand As String = "usp_Sel_SGcm14"
            Dim strTableName As String = "SGcm14"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''	查詢 ((條件：縣市、公私立[1:公立 2:私立])) 園所、班級明細表 公私立 混齡班數量
        Public Function usp_Sel_SGcm15(ByVal strCityID As String, ByVal strGPP As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGcm15"
            Dim strTableName As String = "SGcm15"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GPP", strGPP)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢 縣市、巡撫教授、巡輔員、受輔導班級數
        Public Function usp_Sel_SGcm16() As DataSet
            Dim strCommand As String = "usp_Sel_SGcm16"
            Dim strTableName As String = "SGcm16"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢 縣市名稱、縣市編號、園所編號、班級編號、園所名稱、班級名稱、2-5歲人數總和
        Public Function usp_Sel_SGcm17() As DataSet
            Dim strCommand As String = "usp_Sel_SGcm17"
            Dim strTableName As String = "SGcm17"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''		查詢 ((條件：縣市編號、園所編號、班級編號、身分)) GC_Member 資料表、姓名、key、vi
        Public Function usp_Sel_SGcm18(ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String, ByVal strGCIdentity As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGcm18"
            Dim strTableName As String = "SGcm18"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            dbCheck.htParam.Add("@GCIdentity", strGCIdentity)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢 縣市名稱、縣市編號、園所編號、班級編號、園所名稱、班級名稱、2-5歲各別人數
        Public Function usp_Sel_SGcm19() As DataSet
            Dim strCommand As String = "usp_Sel_SGcm19"
            Dim strTableName As String = "SGcm19"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''查詢 (該縣市、該巡輔員) 即時教保輔導人員明細 查詢縣市、幼兒園、班級、身分別、教保服務人員、帳號、巡迴輔導員
        Public Function usp_Sel_SGcm20(ByVal strCityID As String, ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGcm20"
            Dim strTableName As String = "SGcm20"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：縣市、身分、帳號))查詢縣市編號、帳號、身分
        Public Function usp_Sel_SGcm21(ByVal strCityID As String, ByVal strGCIdentity As String, ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGcm21"
            Dim strTableName As String = "SGcm21"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GCIdentity", strGCIdentity)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''巡輔員的下拉式選單 @A01 教授帳號   @GCIdentity 巡輔員身分C
        Public Function usp_Sel_SGcm22(ByVal strA01 As String, ByVal strGCIdentity As String) As DataSet
            Dim strCommand As String = "usp_Sel_SGcm22"
            Dim strTableName As String = "SGcm22"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@GCIdentity", strGCIdentity)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢姓名、Key、VI、帳號
        Public Function usp_Sel_SEcm1(ByVal strCRID As String, ByVal strGCIdentity As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcm1"
            Dim strTableName As String = "SEcm1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@GCIdentity", strGCIdentity)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：帳號))教保輔導人員資料表
        Public Function usp_Sel_SEcm2(ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcm2"
            Dim strTableName As String = "SEcm2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：XXX+1911-09-01、XXX+1911+1-06-30))入班輔導記錄統計表
        Public Function usp_Sel_SEcm3(ByVal stryear1 As String, ByVal stryear2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcm3"
            Dim strTableName As String = "SEcm3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@year1", stryear1)
            dbCheck.htParam.Add("@year2", stryear2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：XXX+1911-09-01、XXX+1911+1-06-30))入班輔導教授回饋明細
        Public Function usp_Sel_SEcm4(ByVal stryear1 As String, ByVal stryear2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcm4"
            Dim strTableName As String = "SEcm4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@year1", stryear1)
            dbCheck.htParam.Add("@year2", stryear2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：XXX+1911-09-01、XXX+1911+1-06-30))教授回饋統計表
        Public Function usp_Sel_SEcm5(ByVal stryear1 As String, ByVal stryear2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcm5"
            Dim strTableName As String = "SEcm5"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@year1", stryear1)
            dbCheck.htParam.Add("@year2", stryear2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：XXX+1911-09-01、XXX+1911+1-06-30))建議追蹤原始資料
        Public Function usp_Sel_SEcm6(ByVal stryear1 As String, ByVal stryear2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcm6"
            Dim strTableName As String = "SEcm6"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@year1", stryear1)
            dbCheck.htParam.Add("@year2", stryear2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢指標版本(有在使用的)
        Public Function usp_Sel_STgt1() As DataSet
            Dim strCommand As String = "usp_Sel_STgt1"
            Dim strTableName As String = "STgt1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢最新指標版本
        Public Function usp_Sel_STgt2(ByVal strTVersion As String, ByVal strTType As String) As DataSet
            Dim strCommand As String = "usp_Sel_STgt2"
            Dim strTableName As String = "STgt2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TVersion", strTVersion)
            dbCheck.htParam.Add("@TType", strTType)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢指標名稱
        Public Function usp_Sel_STgt3(ByVal strTVersion As String, ByVal strTID As String) As DataSet
            Dim strCommand As String = "usp_Sel_STgt3"
            Dim strTableName As String = "STgt3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TVersion", strTVersion)
            dbCheck.htParam.Add("@TID", strTID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢所有指標版本(大到小)
        Public Function usp_Sel_STgt4() As DataSet
            Dim strCommand As String = "usp_Sel_STgt4"
            Dim strTableName As String = "STgt4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢指標編號、指標名稱、指標類型
        Public Function usp_Sel_STgt5(ByVal strTVersion As String) As DataSet
            Dim strCommand As String = "usp_Sel_STgt5"
            Dim strTableName As String = "STgt5"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TVersion", strTVersion)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢最大指標編號(層面)
        Public Function usp_Sel_STgt6(ByVal strTVersion As String) As DataSet
            Dim strCommand As String = "usp_Sel_STgt6"
            Dim strTableName As String = "STgt6"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TVersion", strTVersion)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢指標編號(層面)、指標名稱
        Public Function usp_Sel_STgt7(ByVal strTVersion As String) As DataSet
            Dim strCommand As String = "usp_Sel_STgt7"
            Dim strTableName As String = "STgt7"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TVersion", strTVersion)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢指標編號(向度)
        Public Function usp_Sel_STgt8(ByVal strTVersion As String, ByVal strTID1 As String, ByVal strTID2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_STgt8"
            Dim strTableName As String = "STgt8"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TVersion", strTVersion)
            dbCheck.htParam.Add("@TID1", strTID1)
            dbCheck.htParam.Add("@TID2", strTID2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢指標編號(向度)、指標名稱
        Public Function usp_Sel_STgt9(ByVal strTVersion As String, ByVal strTID1 As String, ByVal strTID2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_STgt9"
            Dim strTableName As String = "STgt9"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TVersion", strTVersion)
            dbCheck.htParam.Add("@TID1", strTID1)
            dbCheck.htParam.Add("@TID2", strTID2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢指標編號(指標)
        Public Function usp_Sel_STgt10(ByVal strTVersion As String, ByVal strTID1 As String, ByVal strTID2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_STgt10"
            Dim strTableName As String = "STgt10"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TVersion", strTVersion)
            dbCheck.htParam.Add("@TID1", strTID1)
            dbCheck.htParam.Add("@TID2", strTID2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢指標版本、指標編號
        Public Function usp_Sel_STgt11(ByVal strTVersion As String, ByVal strTType As String, ByVal strTID2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_STgt11"
            Dim strTableName As String = "STgt11"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TVersion", strTVersion)
            dbCheck.htParam.Add("@TType", strTType)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢最新指標版本(不管有沒有在使用，避免新增衝突)
        Public Function usp_Sel_STgt12() As DataSet
            Dim strCommand As String = "usp_Sel_STgt12"
            Dim strTableName As String = "STgt12"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢追蹤狀況版本(有設定好版本 ISYorNo=1)
        Public Function usp_Sel_STst1() As DataSet
            Dim strCommand As String = "usp_Sel_STst1"
            Dim strTableName As String = "STst1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢追蹤狀況編號、追蹤狀況名稱
        Public Function usp_Sel_STst2(ByVal strTSVersion As String) As DataSet
            Dim strCommand As String = "usp_Sel_STst2"
            Dim strTableName As String = "STst2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TSVersion", strTSVersion)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢所有追蹤狀況版本(大到小)
        Public Function usp_Sel_STst3() As DataSet
            Dim strCommand As String = "usp_Sel_STst3"
            Dim strTableName As String = "STst3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢追蹤狀況編號、追蹤狀況名稱、追蹤狀況定義
        Public Function usp_Sel_STst4(ByVal strTSVersion As String) As DataSet
            Dim strCommand As String = "usp_Sel_STst4"
            Dim strTableName As String = "STst4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TSVersion", strTSVersion)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢最大追蹤狀況編號
        Public Function usp_Sel_STst5(ByVal strTSVersion As String) As DataSet
            Dim strCommand As String = "usp_Sel_STst5"
            Dim strTableName As String = "STst5"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TSVersion", strTSVersion)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢最新追蹤狀況版本(不管有沒有在使用，避免新增衝突)
        Public Function usp_Sel_STst6() As DataSet
            Dim strCommand As String = "usp_Sel_STst6"
            Dim strTableName As String = "STst6"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢(版本、最新追蹤狀況編號)追蹤狀況名稱、追蹤狀況定義
        Public Function usp_Sel_STst7(ByVal strTSVersion As String, ByVal strNewTSID As String) As DataSet
            Dim strCommand As String = "usp_Sel_STst7"
            Dim strTableName As String = "STst7"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TSVersion", strTSVersion)
            dbCheck.htParam.Add("@NewTSID", strNewTSID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢設定好改善原因版本
        Public Function usp_Sel_SIpr1() As DataSet
            Dim strCommand As String = "usp_Sel_SIpr1"
            Dim strTableName As String = "SIpr1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''查詢((條件：版本編號、是否使用Y))改善原因編號、改善原因名稱、改善原因定義
        Public Function usp_Sel_SIpr2(ByVal strIRVersion As String) As DataSet
            Dim strCommand As String = "usp_Sel_SIpr2"
            Dim strTableName As String = "SIpr2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@IRVersion", strIRVersion)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢所有改善原因版本(大到小)
        Public Function usp_Sel_SIpr3() As DataSet
            Dim strCommand As String = "usp_Sel_SIpr3"
            Dim strTableName As String = "SIpr3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：版本))改善原因編號、改善原因名稱、改善原因定義
        Public Function usp_Sel_SIpr4(ByVal strIRVersion As String) As DataSet
            Dim strCommand As String = "usp_Sel_SIpr4"
            Dim strTableName As String = "SIpr4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@IRVersion", strIRVersion)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：版本))最大改善原因編號
        Public Function usp_Sel_SIpr5(ByVal strIRVersion As String) As DataSet
            Dim strCommand As String = "usp_Sel_SIpr5"
            Dim strTableName As String = "SIpr5"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@IRVersion", strIRVersion)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢最新改善原因版本(不管有沒有在使用，避免新增衝突)
        Public Function usp_Sel_SIpr6() As DataSet
            Dim strCommand As String = "usp_Sel_SIpr6"
            Dim strTableName As String = "SIpr6"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢輔導紀錄編號、建議編號、內容、指標版本、指標編號、追蹤狀況版本、追蹤狀況編號、身份、結案日期、入班日期
        Public Function usp_Sel_SEcs1(ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcs1"
            Dim strTableName As String = "SEcs1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢建議編號、內容、指標版本、指標編號、追蹤狀況版本、追蹤狀況編號、身份
        Public Function usp_Sel_SEcs2(ByVal strCRID As String, ByVal strSType As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcs2"
            Dim strTableName As String = "SEcs2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@SType", strSType)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢查詢最大建議編號
        Public Function usp_Sel_SEcs3(ByVal strCRID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcs3"
            Dim strTableName As String = "SEcs3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢教保輔導優點與改善策略建議資料表、教保輔導紀錄資料表、追蹤狀況資料表((條件：縣市、園所、編號))CRID輔導紀錄編號，CRSID建議編號、入班日期、New追蹤狀況編號、追蹤狀態版本、身份、內容
        Public Function usp_Sel_SEcs4(ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcs4"
            Dim strTableName As String = "SEcs4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function



        '''查詢最大輔導紀錄編號
        Public Function usp_Sel_SEcr1(ByVal strYearNum As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcr1"
            Dim strTableName As String = "SEcr1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@YearNum", strYearNum)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢最小入班日期
        Public Function usp_Sel_SEcr2() As DataSet
            Dim strCommand As String = "usp_Sel_SEcr2"
            Dim strTableName As String = "SEcr2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：使用者。自由查詢：民國年、縣市、園所、班級))輔導紀錄編號、入班日期、縣市名稱、園所名稱、班級名稱、狀態
        '''查詢輔導紀錄編號、入班日期、縣市名稱、園所名稱、班級名稱、狀態
        Public Function usp_Sel_SEcr3(ByVal strA01 As String, ByVal strYear As String, ByVal strCityID As String,
                                      ByVal intGID As String, ByVal intGCID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcr3"
            Dim strTableName As String = "SEcr3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", intGID)
            dbCheck.htParam.Add("@GCID", intGCID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        ''''查詢((條件：使用者。自由查詢：民國年、縣市、園所、班級))輔導紀錄編號、入班日期、縣市名稱、園所名稱、班級名稱、狀態
        ''''查詢輔導紀錄編號、入班日期、縣市名稱、園所名稱、班級名稱、狀態
        'Public Function usp_Sel_SEcr3(ByVal strA01 As String) As DataSet
        '    Dim strCommand As String = "usp_Sel_SEcr3"
        '    Dim strTableName As String = "SEcr3"
        '    Dim dbCheck As New PISCP.DB.ClassDB
        '    Dim dsDone As New DataSet()
        '    Dim spParams As SqlParameter() = Nothing

        '    dbCheck.htParam.Clear()
        '    dbCheck.htParam.Add("@A01", strA01)
        '    'dbCheck.htParam.Add("@Years", strYears)
        '    'dbCheck.htParam.Add("@CityID", strCityID)
        '    'dbCheck.htParam.Add("@GID", strGID)
        '    'dbCheck.htParam.Add("@GCID", strGCID)
        '    spParams = dbCheck.getSpParams(0)
        '    dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

        '    Return dsDone
        'End Function

        '''查詢入班日期、結案日期、縣市編號、園所編號、班級編號、指標版本、指標編號、開始時間、結束時間、開始時間(含交通)、結束時間(含交通)、是否連續駐區5天、應到人數、應到人數(5歲幼兒)、實到人數、教學名稱、幼兒健康狀況說明、園所處理狀況說明、入班觀察及輔導流程、觀察或輔導事件描述、資深巡輔員回饋、輔導教授回饋、狀態
        Public Function usp_Sel_SEcr4(ByVal strCRID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcr4"
            Dim strTableName As String = "SEcr4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢最大入班日期、開始時間、結束時間、開始時間(含交通)、結束時間(含交通)、指標編號、應到人數、應到人數(5歲幼兒)、幼兒健康狀況說明、園所處理狀況說明
        Public Function usp_Sel_SEcr5(ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcr5"
            Dim strTableName As String = "SEcr5"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢月輔導報告-非教保輔導資料表((條件：巡輔員帳號、西元年int、月份int、縣市))報告日期、項目編號、日期、開始時間、結束時間、機構名稱、處理業務簡述、備註
        Public Function usp_Sel_SEcr6(ByVal strA01 As String, ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcr6"
            Dim strTableName As String = "SEcr6"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢((條件：帳號，完稿2))計算該學期EduCR總比數用，用count
        Public Function usp_Sel_SEcr7(ByVal strA01 As String, ByVal strYeardate1 As String, ByVal strYeardate2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcr7"
            Dim strTableName As String = "SEcr7"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@Yeardate1", strYeardate1)
            dbCheck.htParam.Add("@Yeardate2", strYeardate2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''查詢幼兒園是否有 教保輔導紀錄 ((條件：縣市、班級)) 最大入班日期、開始時間、結束時間、開始時間(含交通)、結束時間(含交通)、應到人數、應到人數(5歲幼兒)、幼兒健康狀況說明、園所處理狀況說明
        Public Function usp_Sel_SEcr8(ByVal strCityID As String, ByVal strGID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcr8"
            Dim strTableName As String = "SEcr8"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢最大入班日期、開始時間、結束時間、開始時間(含交通)、結束時間(含交通)、指標編號、應到人數、應到人數(5歲幼兒)、幼兒健康狀況說明、園所處理狀況說明
        Public Function usp_Sel_SEcr9(ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String, ByVal strCRID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcr9"
            Dim strTableName As String = "SEcr9"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            dbCheck.htParam.Add("@CRID", strCRID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢症狀、人數
        Public Function usp_Sel_SEch1(ByVal strCRID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEch1"
            Dim strTableName As String = "SEch1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢(舊資料)前次建議追蹤(前次輔導紀錄編號、前次建議編號、內容、指標版本、指標編號、追蹤狀況版本、追蹤狀況編號、身份、結案日期、結案內容)
        Public Function usp_Sel_SOld1(ByVal strCRID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SOld1"
            Dim strTableName As String = "SOld1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢(舊資料)前次建議追蹤(條件：前次輔導紀錄編號、前次建議編號)(輔導紀錄編號、前次輔導紀錄編號、前次建議編號、內容、指標版本、指標編號、追蹤狀況版本、追蹤狀況編號、身份、結案日期、結案內容、入班日期)
        Public Function usp_Sel_SOld2(ByVal strOldCRID As String, ByVal strOldCRSID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SOld2"
            Dim strTableName As String = "SOld2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@OldCRID", strOldCRID)
            dbCheck.htParam.Add("@OldCRSID", strOldCRSID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''查詢公告編號、公告時間、主題、內容
        Public Function usp_Sel_SNtF1() As DataSet
            Dim strCommand As String = "usp_Sel_SNtF1"
            Dim strTableName As String = "SNtF1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢公告日期、主題、內容
        Public Function usp_Sel_SNtF2(ByVal strNFID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SNtF2"
            Dim strTableName As String = "SNtF2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@NFID", strNFID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢最大公告編號
        Public Function usp_Sel_SNtF3() As DataSet
            Dim strCommand As String = "usp_Sel_SNtF3"
            Dim strTableName As String = "SNtF3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢公告日期區間
        Public Function usp_Sel_SNtF4(ByVal strTimes1 As String, ByVal strTimes2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SNtF4"
            Dim strTableName As String = "SNtF4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@Times1", strTimes1)
            dbCheck.htParam.Add("@Times2", strTimes2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''查詢最大問卷編號
        Public Function usp_Sel_SQme1() As DataSet
            Dim strCommand As String = "usp_Sel_SQme1"
            Dim strTableName As String = "SQme1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷管理資料表(問卷編號、問卷名稱、問卷說明、開放填答開始時間、開放填答最後時間)
        Public Function usp_Sel_SQme2(ByVal strQID As String, ByVal strQName As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQme2"
            Dim strTableName As String = "SQme2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QName", strQName)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷管理資料表(((條件：問卷編號))問卷名稱、問卷說明、開放填答開始時間、開放填答最後時間)
        Public Function usp_Sel_SQme3(ByVal strQID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQme3"
            Dim strTableName As String = "SQme3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷管理資料表((限定該受訪人員、自由查詢條件：問卷編號、問卷名稱))問卷編號、問卷名稱、問卷說明、開放填答開始時間、開放填答最後時間
        Public Function usp_Sel_SQme4(ByVal strA01 As String, ByVal strQID As String, ByVal strQName As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQme4"
            Dim strTableName As String = "SQme4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QName", strQName)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢最大題目分類編號((條件：問卷編號))
        Public Function usp_Sel_SQcf1(ByVal strQID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQcf1"
            Dim strTableName As String = "SQcf1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷題目分類(((條件：問卷編號))題目分類編號、題目分類名稱、題目分類說明)
        Public Function usp_Sel_SQcf2(ByVal strQID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQcf2"
            Dim strTableName As String = "SQcf2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷題目分類第一筆(((條件：問卷編號))題目分類編號、題目分類名稱、題目分類說明)
        Public Function usp_Sel_SQcf3(ByVal strQID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQcf3"
            Dim strTableName As String = "SQcf3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷題目分類名稱(((條件：問卷編號、題目分類編號))題目分類編號、題目分類名稱、題目分類說明)
        Public Function usp_Sel_SQcf4(ByVal strQID As String, ByVal strQCID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQcf4"
            Dim strTableName As String = "SQcf4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢最大題目編號((條件：問卷編號、題目分類編號))
        Public Function usp_Sel_SQue1(ByVal strQID As String, ByVal strQCID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQue1"
            Dim strTableName As String = "SQue1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷題目((條件：問卷編號、題目分類編號)) 問卷編號、題目分類編號、題目編號、題目名稱、選項類型、必填、其他
        Public Function usp_Sel_SQue2(ByVal strQID As String, ByVal strQCID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQue2"
            Dim strTableName As String = "SQue2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷題目((條件：問卷編號、題目分類編號、題目編號)) 問卷編號、題目分類編號、題目編號、題目名稱、選項類型、必填、其他
        Public Function usp_Sel_SQue3(ByVal strQID As String, ByVal strQCID As String, ByVal strQTID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQue3"
            Dim strTableName As String = "SQue3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            dbCheck.htParam.Add("@QTID", strQTID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷題目-下一頁用((條件：問卷編號、題目分類編號、題目編號)) 問卷編號、題目分類編號、題目編號、題目名稱、選項類型、必填、其他
        Public Function usp_Sel_SQue4(ByVal strQID As String, ByVal strQCID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQue4"
            Dim strTableName As String = "SQue4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷題目跳題用((條件：問卷編號、題目編號)) 問卷編號、題目分類編號、題目分類名稱、題目編號、題目名稱
        Public Function usp_Sel_SQue5(ByVal strQID As String, ByVal strQTID As String, ByVal strQTtype As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQue5"
            Dim strTableName As String = "SQue5"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QTID", strQTID)
            dbCheck.htParam.Add("@QTtype", strQTtype)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢最大題目編號((條件：問卷編號))
        Public Function usp_Sel_SQue6(ByVal strQID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQue6"
            Dim strTableName As String = "SQue6"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷題目-上一頁用((條件：問卷編號、題目分類編號、題目編號)) 問卷編號、題目分類編號、題目編號、題目名稱、選項類型、必填、其他
        Public Function usp_Sel_SQue7(ByVal strQID As String, ByVal strQCID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQue7"
            Dim strTableName As String = "SQue7"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷題目-最大題目編號((條件：問卷編號)) 題目分類編號
        Public Function usp_Sel_SQue9(ByVal strQID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQue9"
            Dim strTableName As String = "SQue9"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷題目-最小題目編號((條件：問卷編號)) 題目分類編號
        Public Function usp_Sel_SQue10(ByVal strQID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQue10"
            Dim strTableName As String = "SQue10"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷題目((條件：問卷編號)) 題目分類編號、題目編號、題目名稱、選項類型、必填、其他
        Public Function usp_Sel_SQue11(ByVal strQID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQue11"
            Dim strTableName As String = "SQue11"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷受訪人員((條件：問卷編號、帳號)) 問卷編號、帳號、填答時間、狀態
        Public Function usp_Sel_SQur1(ByVal strQID As String, ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQur1"
            Dim strTableName As String = "SQur1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷受訪人員連動Acc資料表解析名字((條件：問卷編號)) 問卷編號、帳號、填答時間、狀態、姓名、使用者身份、Key、VI
        Public Function usp_Sel_SQur2(ByVal strQID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQur2"
            Dim strTableName As String = "SQur2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷受訪人員連動Acc資料表((條件：問卷編號、使用者身份)) 問卷編號、帳號、填答時間、狀態、姓名、使用者身份、Key、VI
        Public Function usp_Sel_SQur3(ByVal strQID As String, ByVal strA19 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQur3"
            Dim strTableName As String = "SQur3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@A19", strA19)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''查詢問卷受訪人員連動Acc資料表((條件：上學期時間9/1，1/31)) 縣市名稱、園所名稱、帳號、姓名、Key、VI、Email、權限
        Public Function usp_Sel_SQur4(ByVal stryear1 As String, ByVal stryear2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQur4"
            Dim strTableName As String = "SQur4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@year1", stryear1)
            dbCheck.htParam.Add("@year2", stryear2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        ''' 查詢問卷受訪人員計算((條件：上學期時間9/1，1/31，縣市，填答狀態1、填答狀態2。填答狀態 0 未完成 1 完成 其他內容 皆是會對不到來判斷總人數與填答人數用 ))Count QuestUser.A01 數量
        Public Function usp_Sel_SQur5(ByVal stryear1 As String, ByVal stryear2 As String, ByVal strCityID As String, ByVal strQState1 As String,
                                      ByVal strQState2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQur5"
            Dim strTableName As String = "SQur5"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@year1", stryear1)
            dbCheck.htParam.Add("@year2", stryear2)
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@QState1", strQState1)
            dbCheck.htParam.Add("@QState2", strQState2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷受訪人員連動Acc資料表((條件：上學期時間9/1，1/31)) 縣市名稱、園所名稱、帳號、姓名、Key、VI、填答完成時間、IP位置
        Public Function usp_Sel_SQur6(ByVal stryear1 As String, ByVal stryear2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQur6"
            Dim strTableName As String = "SQur6"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@year1", stryear1)
            dbCheck.htParam.Add("@year2", stryear2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢該學期問卷題目((條件：上學期時間9/1，1/31)) 題目名稱、題目分類編號、題目編號
        Public Function usp_Sel_SQur7(ByVal stryear1 As String, ByVal stryear2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQur7"
            Dim strTableName As String = "SQur7"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@year1", stryear1)
            dbCheck.htParam.Add("@year2", stryear2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢該學期問卷題目教保員未填答名單((條件：上學期時間9/1，1/31、縣市)) 園所名稱、帳號、姓名
        Public Function usp_Sel_SQur8(ByVal stryear1 As String, ByVal stryear2 As String, ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQur8"
            Dim strTableName As String = "SQur8"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@year1", stryear1)
            dbCheck.htParam.Add("@year2", stryear2)
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢該學期問卷題目((條件：問卷編號)) 
        Public Function usp_Sel_SQur9(ByVal stryQID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQur9"
            Dim strTableName As String = "SQur9"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", stryQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢該學期問卷填答完成者與IP((條件：問卷編號)) 
        Public Function usp_Sel_SQur10(ByVal stryQID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQur10"
            Dim strTableName As String = "SQur10"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", stryQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢該學期問卷填答內容((條件：問卷編號、填答帳號)) 
        Public Function usp_Sel_SQur11(ByVal stryQID As String, ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQur11"
            Dim strTableName As String = "SQur11"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", stryQID)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢最大選項編號((條件：問卷編號、題目分類編號、題目編號))
        Public Function usp_Sel_SQop1(ByVal strQID As String, ByVal strQCID As String, ByVal strQTID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQop1"
            Dim strTableName As String = "SQop1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            dbCheck.htParam.Add("@QTID", strQTID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：問卷編號))題目分類名稱、題目分類編號、選項編號、必填、題目名稱、題目編號
        Public Function usp_Sel_SQop2(ByVal strQID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQop2"
            Dim strTableName As String = "SQop2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：問卷編號、題目分類編號、題目編號))選項編號、選項名稱
        Public Function usp_Sel_SQop3(ByVal strQID As String, ByVal strQCID As String, ByVal strQTID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQop3"
            Dim strTableName As String = "SQop3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            dbCheck.htParam.Add("@QTID", strQTID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：問卷編號、題目分類編號、題目編號、選項編號))選項編號、選項名稱
        Public Function usp_Sel_SQop4(ByVal strQID As String, ByVal strQCID As String, ByVal strQTID As String, ByVal strQOID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQop4"
            Dim strTableName As String = "SQop4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            dbCheck.htParam.Add("@QTID", strQTID)
            dbCheck.htParam.Add("@QOID", strQOID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''查詢((條件：問卷編號、題目分類編號))題目分類名稱、題目分類編號、題目編號、題目名稱、必填、其他、選項編號、選項名稱、跳題
        Public Function usp_Sel_SQop5(ByVal strQID As String, ByVal strQCID As String, ByVal strQTID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQop5"
            Dim strTableName As String = "SQop5"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            dbCheck.htParam.Add("@QTID", strQTID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢最大選項編號不包含999((條件：問卷編號、題目分類編號、題目編號、<>999))
        Public Function usp_Sel_SQop6(ByVal strQID As String, ByVal strQCID As String, ByVal strQTID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQop6"
            Dim strTableName As String = "SQop6"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            dbCheck.htParam.Add("@QTID", strQTID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢選項列表，不包含999((條件：問卷編號、題目分類編號、題目編號、<>999))選項編號、選項名稱
        Public Function usp_Sel_SQop7(ByVal strQID As String, ByVal strQCID As String, ByVal strQTID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQop7"
            Dim strTableName As String = "SQop7"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            dbCheck.htParam.Add("@QTID", strQTID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢選項列表，不包含999((條件：問卷編號、題目編號、<>999))選項編號、選項名稱、跳題
        Public Function usp_Sel_SQop8(ByVal strQID As String, ByVal strQTID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQop8"
            Dim strTableName As String = "SQop8"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QTID", strQTID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢選項資料表((條件：問卷編號、跳題編號(此為刪除的那題編號)))題目編號、跳題
        Public Function usp_Sel_SQop9(ByVal strQID As String, ByVal strQOjump As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQop9"
            Dim strTableName As String = "SQop9"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QOjump", strQOjump)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        ''' 查詢問卷選項資料表((條件：問卷編號))題目編號、選項編號、題目分類編號、選項名稱、跳題
        Public Function usp_Sel_SQop10(ByVal strQID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQop10"
            Dim strTableName As String = "SQop10"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢((條件：問卷編號、題目分類編號、題目編號、選項編號))選項編號、選項名稱
        Public Function usp_Sel_SQrc1(ByVal strQID As String, ByVal strQCID As String, ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQrc1"
            Dim strTableName As String = "SQrc1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷填答紀錄((條件：問卷編號、題目編號、帳號))題目編號、選項編號、其他文字
        Public Function usp_Sel_SQrc2(ByVal strQID As String, ByVal strQTID As String, ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQrc2"
            Dim strTableName As String = "SQrc2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QTID", strQTID)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷填答是否有紀錄((條件：問卷編號))選項編號
        Public Function usp_Sel_SQrc3(ByVal strQID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQrc3"
            Dim strTableName As String = "SQrc3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷填答者填答未完成數量((條件：帳號，學期時間起迄))問卷編號、縣市名稱
        Public Function usp_Sel_SQrc4(ByVal strA01 As String, ByVal strYeardate1 As String, ByVal strYeardate2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQrc4"
            Dim strTableName As String = "SQrc4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@Yeardate1", strYeardate1)
            dbCheck.htParam.Add("@Yeardate2", strYeardate2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢問卷填答者填答完成數量((條件：帳號，學期時間起迄))問卷編號、縣市名稱
        Public Function usp_Sel_SQrc5(ByVal strA01 As String, ByVal strYeardate1 As String, ByVal strYeardate2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SQrc5"
            Dim strTableName As String = "SQrc5"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@Yeardate1", strYeardate1)
            dbCheck.htParam.Add("@Yeardate2", strYeardate2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢月輔導報告資料表((條件：巡輔員帳號、西元年int、月份int、縣市))應上班天數、應輔導天數、教保輔導天數、非教保輔導天數、需要協助或建議、縣市政府之相關意見及處理策略
        Public Function usp_Sel_SRecrm1(ByVal strA01 As String, ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SRecrm1"
            Dim strTableName As String = "SRecrm1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢月輔導報告-非教保輔導資料表((條件：巡輔員帳號、西元年int、月份int、縣市))報告日期、項目編號、日期、開始時間、結束時間、機構名稱、處理業務簡述、備註
        Public Function usp_Sel_SRnecrm1(ByVal strA01 As String, ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SRnecrm1"
            Dim strTableName As String = "SRnecrm1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢月輔導報告-非教保輔導資料表((條件：西元年int、月份int、縣市))最大項目編號
        Public Function usp_Sel_SRnecrm2(ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SRnecrm2"
            Dim strTableName As String = "SRnecrm2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢輔導教授每月填報意見資料表((條件：巡輔員帳號、西元年int、月份int、縣市))需要協助或建議、縣市政府之相關意見及處理策略
        Public Function usp_Sel_SReportPA1(ByVal strA01 As String, ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SReportPA1"
            Dim strTableName As String = "SReportPA1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢輔導教授每月填報園所資料表((條件：輔導教授帳號、西元年int、月份int、縣市))報告日期、項目編號、日期、園所編號、內容(訪視心得/省思)
        Public Function usp_Sel_SReportPG1(ByVal strA01 As String, ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SReportPG1"
            Dim strTableName As String = "SReportPG1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢輔導教授每月填報園所資料表((條件：西元年int、月份int、縣市))最大項目編號
        Public Function usp_Sel_SReportPG2(ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SReportPG2"
            Dim strTableName As String = "SReportPG2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢輔導教授每月填報園所資料表((條件：西元年int、月份int、縣市))最大項目編號
        Public Function usp_Sel_SReportPG3(ByVal strCityID As String, ByVal strGID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SReportPG3"
            Dim strTableName As String = "SReportPG3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢巡迴輔導教授時數((條件：XXX+1911-09-01、XXX+1911+1-06-30))姓名、key、Vi、日期、園所/會議名稱、天數 (查詢輔導教授每月紀錄算1天、巡輔員填寫的會議紀錄算0.5天
        Public Function usp_Sel_SReportPG4(ByVal stryear1 As String, ByVal stryear2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SReportPG4"
            Dim strTableName As String = "SReportPG4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@year1", stryear1)
            dbCheck.htParam.Add("@year2", stryear2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢每月上傳檔案資料表((條件：西元年int、月份int、縣市))最大檔案編號
        Public Function usp_Sel_SUlf1(ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SUlf1"
            Dim strTableName As String = "SUlf1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''	查詢每月上傳檔案資料表((條件：巡輔員帳號、西元年int、月份int、縣市))檔案編號、類型、檔案名稱、副檔名
        Public Function usp_Sel_SUlf2(ByVal strA01 As String, ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SUlf2"
            Dim strTableName As String = "SUlf2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢每月上傳檔案((條件：西元年int、月份int、縣市、檔案編號))類型、檔案名稱、副檔名
        Public Function usp_Sel_SUlf3(ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String, ByVal strFileID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SUlf3"
            Dim strTableName As String = "SUlf3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@FileID", strFileID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢每月上傳檔案資料表((條件：西元年int、月份int、縣市))最大檔案編號
        Public Function usp_Sel_SAmt1(ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAmt1"
            Dim strTableName As String = "SAmt1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢每月會議記錄出席教授資料表((條件：巡輔員帳號、、西元年int、月份int、縣市))編號、帳號、日期、會議名稱
        Public Function usp_Sel_SAmt2(ByVal strAA01 As String, ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAmt2"
            Dim strTableName As String = "SAmt2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@AA01", strAA01)
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        ''' 查詢前次建議追蹤(
        ''' 1. 教保輔導前次建議追蹤資料表 OldCRID前次輔導紀錄編號，OldCRSID前次建議編號
        ''' 2. 教保輔導優點與改善策略建議資料表： Suggest 內容， TVersion 指標版本，TID 指標編號， TSVersion 追蹤狀況版本 ，TSID 追蹤狀況編號，GCIdentity 身分
        ''' 3. 教保輔導紀錄資料表 ：CityID 縣市編號 、GID 園所編號、GCID 班級編號、EntryDate入班日期
        ''' 4. 行政輔導紀錄資料表 ： TSID 追蹤狀況編號、 Contents 內容、ReportDate 回報日期、TrackWay 追蹤方式)
        Public Function usp_Sel_SAmr1(ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String,
                                      ByVal strReportDate As String, ByVal strTrackWay As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAmr1"
            Dim strTableName As String = "SAmr1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@TrackWay", strTrackWay)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '查詢行政輔導紀錄資料表((條件：回報日期、前次輔導紀錄編號、前次建議編號、縣市、幼兒園、班級、追蹤方式))追蹤狀況編號、追蹤狀態、內容
        Public Function usp_Sel_SAmr2(ByVal strReportDate As String, ByVal strOldCRID As String, ByVal strOldCRSID As String,
                                      ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String, ByVal strTrackWay As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAmr2"
            Dim strTableName As String = "SAmr2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@OldCRID", strOldCRID)
            dbCheck.htParam.Add("@OldCRSID", strOldCRSID)
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            dbCheck.htParam.Add("@TrackWay", strTrackWay)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '查詢行政輔導紀錄資料表((條件：回報日期、縣市、幼兒園、班級、追蹤方式))回報日期、縣市編號、園所編號、班級編號、追蹤方式、前次輔導紀錄編號、前次建議編號、追蹤狀況編號、內容
        Public Function usp_Sel_SAmr3(ByVal strReportDate As String, ByVal strCityID As String, ByVal strGID As String,
                                      ByVal strGCID As String, ByVal strTrackWay As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAmr3"
            Dim strTableName As String = "SAmr3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            dbCheck.htParam.Add("@TrackWay", strTrackWay)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '查詢行政輔導紀錄資料表((條件：回報日期區間如'1911-01-01'~'1911-02-28'、縣市、幼兒園、班級))回報日期、縣市編號、園所編號、班級編號、追蹤方式
        Public Function usp_Sel_SAmr4(ByVal strReportDate1 As String, ByVal strReportDate2 As String, ByVal strCityID As String,
                                      ByVal strGID As String, ByVal strGCID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAmr4"
            Dim strTableName As String = "SAmr4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@ReportDate1", strReportDate1)
            dbCheck.htParam.Add("@ReportDate2", strReportDate2)
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function



        '查詢行政輔導紀錄資料表((條件：回報日期、縣市、幼兒園、班級))回報日期、縣市編號、園所編號、班級編號、追蹤方式、前次輔導紀錄編號、前次建議編號、追蹤狀況編號、內容
        Public Function usp_Sel_SAmr5(ByVal strReportDate As String, ByVal strCityID As String, ByVal strGID As String,
                                      ByVal strGCID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAmr5"
            Dim strTableName As String = "SAmr5"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '查詢行政輔導紀錄資料表，刪除前復原用，只查前兩筆((條件：前次輔導紀錄編號、前次建議編號))回報日期、縣市編號、園所編號、班級編號、追蹤方式、前次輔導紀錄編號、前次建議編號、追蹤狀況編號、內容
        Public Function usp_Sel_SAmr6(ByVal strOldCRID As String, ByVal strOldCRSID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAmr6"
            Dim strTableName As String = "SAmr6"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@OldCRID", strOldCRID)
            dbCheck.htParam.Add("@OldCRSID", strOldCRSID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '查詢行政輔導紀錄資料表，刪除檢查用((條件：前次輔導紀錄編號))回報日期、縣市編號、園所編號、班級編號、追蹤方式、前次輔導紀錄編號、前次建議編號、追蹤狀況編號、內容
        Public Function usp_Sel_SAmr7(ByVal strOldCRID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAmr7"
            Dim strTableName As String = "SAmr7"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@OldCRID", strOldCRID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''	查詢((條件：帳號，完稿2))計算該學期EduCR總比數用，用count
        Public Function usp_Sel_SAmr8(ByVal strA01 As String, ByVal strYeardate1 As String, ByVal strYeardate2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAmr8"
            Dim strTableName As String = "SAmr8"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@Yeardate1", strYeardate1)
            dbCheck.htParam.Add("@Yeardate2", strYeardate2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '查詢非入班追蹤紀錄資料表((條件：縣市，園所編號，班級編號))進行檢查比對是否有資料，有資料無法刪除
        Public Function usp_Sel_SAmr9(ByVal strCityID As String, ByVal strGID As String,
                                      ByVal strGCID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAmr9"
            Dim strTableName As String = "SAmr9"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '查詢非入班追蹤紀錄資料表((條件：縣市，園所編號))進行檢查比對是否有資料，有資料無法刪除
        Public Function usp_Sel_SAmr10(ByVal strCityID As String, ByVal strGID As String) As DataSet
            Dim strCommand As String = "usp_Sel_SAmr10"
            Dim strTableName As String = "SAmr10"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''修改帳戶資料(密碼)
        Public Function usp_Upd_UAcc1(ByVal strCode As String, ByVal strPass As String, ByVal strA01 As String) As String
            '查詢帳號、姓名、生日、性別、手機號碼、聯絡電話、電子信箱、教師身份、園所代碼、註冊日期、權限、Key、VI
            Dim strCommand1 As String = "usp_Sel_SAcc1"
            Dim strTableName1 As String = "SAcc1"
            Dim dbCheck1 As New PISCP.DB.ClassDB
            Dim dsDone1 As New DataSet()
            Dim spParams1 As SqlParameter() = Nothing

            dbCheck1.htParam.Clear()
            dbCheck1.htParam.Add("@A01", strA01)
            spParams1 = dbCheck1.getSpParams(0)
            dsDone1 = dbCheck1.RunProcedure(strCommand1, spParams1, strTableName1)

            If dsDone1.Tables(0).Rows.Count <> 0 Then
                Dim A03, A04, A05, A06, A08, A09, A10 As String
                A05 = AES.AES_Decrypt(dsDone1.Tables(0).Rows(0).Item(1).ToString(), dsDone1.Tables(0).Rows(0).Item(11).ToString(), dsDone1.Tables(0).Rows(0).Item(12).ToString())
                A06 = AES.AES_Decrypt(dsDone1.Tables(0).Rows(0).Item(2).ToString(), dsDone1.Tables(0).Rows(0).Item(11).ToString(), dsDone1.Tables(0).Rows(0).Item(12).ToString())
                A08 = AES.AES_Decrypt(dsDone1.Tables(0).Rows(0).Item(4).ToString(), dsDone1.Tables(0).Rows(0).Item(11).ToString(), dsDone1.Tables(0).Rows(0).Item(12).ToString())
                A09 = AES.AES_Decrypt(dsDone1.Tables(0).Rows(0).Item(5).ToString(), dsDone1.Tables(0).Rows(0).Item(11).ToString(), dsDone1.Tables(0).Rows(0).Item(12).ToString())
                A10 = dsDone1.Tables(0).Rows(0).Item(6).ToString()

                Dim strCommand2 As String = "usp_Upd_UAcc1"
                Dim strTableName2 As String = "UAcc1"
                Dim dbCheck2 As New PISCP.DB.ClassDB
                Dim spParams2 As SqlParameter() = Nothing

                '加密
                A03 = AES.AES_Random(32)
                A04 = AES.AES_Random(16)
                dbCheck2.htParam.Clear()
                dbCheck2.htParam.Add("@A01", strA01)
                dbCheck2.htParam.Add("@A02", AES.AES_Encrypt(strPass, A03, A04))
                dbCheck2.htParam.Add("@A03", A03)
                dbCheck2.htParam.Add("@A04", A04)
                dbCheck2.htParam.Add("@A05", AES.AES_Encrypt(A05, A03, A04))
                dbCheck2.htParam.Add("@A06", AES.AES_Encrypt(A06, A03, A04))
                dbCheck2.htParam.Add("@A08", AES.AES_Encrypt(A08, A03, A04))
                dbCheck2.htParam.Add("@A09", AES.AES_Encrypt(A09, A03, A04))
                dbCheck2.htParam.Add("@A10", A10)
                spParams2 = dbCheck2.getSpParams(0)
                dbCheck2.RunProcedure(strCommand2, spParams2, strTableName2)

                '寄送電子郵件
                '寄件者信箱帳號、*收件者帳號*
                Dim mailMsg1 As MailMessage = New MailMessage("PISCP2020@gmail.com", dsDone1.Tables(0).Rows(0).Item(6).ToString())
                Dim mailSmtp1 As SmtpClient = New SmtpClient("smtp.gmail.com", 25)

                '判斷修改密碼、忘記密碼
                If strCode = "1" Then
                    '主旨
                    mailMsg1.Subject = "【國民教育幼兒班教保輔導系統】修改密碼"
                    '以Html方式發送
                    mailMsg1.IsBodyHtml = True
                    Dim mailstr1 As String = ""
                    mailstr1 = "您的密碼修改成功！"
                    mailMsg1.Body = mailstr1.ToString()
                    '設定帳號、密碼
                    mailSmtp1.Credentials = New System.Net.NetworkCredential("PISCP2020@gmail.com", "NEWPISCP2020")
                    '開啟SSL驗證
                    mailSmtp1.EnableSsl = True
                    '寄信
                    mailSmtp1.Send(mailMsg1)
                    mailMsg1.Dispose()
                Else
                    '主旨
                    mailMsg1.Subject = "【國民教育幼兒班教保輔導系統】忘記密碼"
                    '以Html方式發送
                    mailMsg1.IsBodyHtml = True
                    Dim mailstr1 As String = ""
                    mailstr1 = "您的密碼重新設為「" & strPass & "」。<br/>*小叮嚀：為了您的帳戶安全，登入後請更改密碼、並牢記，感謝您的使用~"
                    mailMsg1.Body = mailstr1.ToString()
                    '設定帳號、密碼
                    mailSmtp1.Credentials = New System.Net.NetworkCredential("PISCP2020@gmail.com", "NEWPISCP2020")
                    '開啟SSL驗證
                    mailSmtp1.EnableSsl = True
                    '寄信
                    mailSmtp1.Send(mailMsg1)
                    mailMsg1.Dispose()
                End If

                Return "OK&&&新密碼已寄到您的信箱~&&&密碼修改成功！&&&"
            Else
                Return "Null&&&查無此帳號！&&&"
            End If
        End Function

        '''查詢KPI((條件：A14身分、KPI項目))標準值
        Public Function usp_Sel_SKPI1(ByVal strA14 As String, ByVal strKPIproject As String) As DataSet
            Dim strCommand As String = "usp_Sel_SKPI1"
            Dim strTableName As String = "SKPI1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A14", strA14)
            dbCheck.htParam.Add("@KPIproject", strKPIproject)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''修改姓名、生日、性別、手機號碼、聯絡電話、電子信箱
        Public Function usp_Upd_UAcc4(ByVal strA05 As String, ByVal strA06 As String, ByVal strA07 As String, ByVal strA08 As String, ByVal strA09 As String, ByVal strA10 As String, ByVal strA01 As String) As DataSet
            '先查Key VI

            Dim strCommand As String = "usp_Upd_UAcc4"
            Dim strTableName As String = "UAcc4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A05", strA05)
            dbCheck.htParam.Add("@A06", strA06)
            dbCheck.htParam.Add("@A07", strA07)
            dbCheck.htParam.Add("@A08", strA08)
            dbCheck.htParam.Add("@A09", strA09)
            dbCheck.htParam.Add("@A10", strA10)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改姓名、手機號碼、連絡電話、電子信箱、幼兒園代碼、教師身份、使用者身份、停用訊息、備註、權限、服務單位
        Public Function usp_Upd_UAcc5(ByVal strA01 As String, ByVal strA05 As String, ByVal strA08 As String, ByVal strA09 As String,
                                        ByVal strA10 As String, ByVal strA12 As String, ByVal strA11 As String, ByVal strA14 As String,
                                      ByVal strA16 As String, ByVal strA17 As String, ByVal strA19 As String, ByVal strA20 As String) As DataSet
            Dim strCommand As String = "usp_Upd_UAcc5"
            Dim strTableName As String = "UAcc5"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@A05", strA05)
            dbCheck.htParam.Add("@A08", strA08)
            dbCheck.htParam.Add("@A09", strA09)
            dbCheck.htParam.Add("@A10", strA10)
            dbCheck.htParam.Add("@A12", strA12)
            dbCheck.htParam.Add("@A11", strA11)
            dbCheck.htParam.Add("@A14", strA14)
            dbCheck.htParam.Add("@A16", strA16)
            dbCheck.htParam.Add("@A17", strA17)
            dbCheck.htParam.Add("@A19", strA19)
            dbCheck.htParam.Add("@A20", strA20)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''修改停用帳戶代號、權限
        Public Function usp_Upd_UAcc6(ByVal strA01 As String, ByVal strA16 As String, ByVal strA19 As String) As DataSet
            Dim strCommand As String = "usp_Upd_UAcc6"
            Dim strTableName As String = "UAcc6"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@A16", strA16)
            dbCheck.htParam.Add("@A19", strA19)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改巡輔員 專任、兼任
        Public Function usp_Upd_UAcc7(ByVal strA01 As String, ByVal strA21 As String) As DataSet
            Dim strCommand As String = "usp_Upd_UAcc7"
            Dim strTableName As String = "UAcc7"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@A21", strA21)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改最後登入時間
        Public Function usp_Upd_UAcc8(ByVal strA01 As String, ByVal strA22 As String) As DataSet
            Dim strCommand As String = "usp_Upd_UAcc8"
            Dim strTableName As String = "UAcc8"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@A22", strA22)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改入班日期、縣市編號、園所編號、班級編號
        Public Function usp_Upd_UEcr1(ByVal strCRID As String, ByVal strEntryDate As String, ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String) As DataSet
            '先查Key VI

            Dim strCommand As String = "usp_Upd_UEcr1"
            Dim strTableName As String = "UEcr1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@EntryDate", strEntryDate)
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改入指標編號(層面)、開始時間、結束時間、開始時間(含交通、資料整理)、結束時間(含交通、資料整理)、是否連續駐區5天、應到人數、5歲幼兒應到人數、實到人數、教學名稱、教保輔導重點
        Public Function usp_Upd_UEcr2(ByVal strCRID As String, ByVal strTID As String, ByVal strStartTime1 As String,
                                      ByVal strEndTime1 As String, ByVal strStartTime2 As String, ByVal strEndTime2 As String,
                                      ByVal strCS As String, ByVal strEstNum1 As String, ByVal strEstNum2 As String,
                                      ByVal strActualNum As String, ByVal strTopic As String, ByVal strEFocus As String) As DataSet

            Dim strCommand As String = "usp_Upd_UEcr2"
            Dim strTableName As String = "UEcr2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@TID ", strTID)
            dbCheck.htParam.Add("@StartTime1", strStartTime1)
            dbCheck.htParam.Add("@EndTime1", strEndTime1)
            dbCheck.htParam.Add("@StartTime2", strStartTime2)
            dbCheck.htParam.Add("@EndTime2", strEndTime2)
            dbCheck.htParam.Add("@CS", strCS)
            dbCheck.htParam.Add("@EstNum1", strEstNum1)
            dbCheck.htParam.Add("@EstNum2", strEstNum2)
            dbCheck.htParam.Add("@ActualNum", strActualNum)
            dbCheck.htParam.Add("@Topic", strTopic)
            dbCheck.htParam.Add("@EFocus", strEFocus)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改幼兒健康狀況說明、園所處理狀況說明
        Public Function usp_Upd_UEcr3(ByVal strCRID As String, ByVal strHealth1 As String, ByVal strHealth2 As String) As DataSet

            Dim strCommand As String = "usp_Upd_UEcr3"
            Dim strTableName As String = "UEcr3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@Health1", strHealth1)
            dbCheck.htParam.Add("@Health2", strHealth2)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改入班觀察及輔導流程
        Public Function usp_Upd_UEcr4(ByVal strCRID As String, ByVal strProcess As String) As DataSet

            Dim strCommand As String = "usp_Upd_UEcr4"
            Dim strTableName As String = "UEcr4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@Process", strProcess)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改觀察或輔導事件描述
        Public Function usp_Upd_UEcr5(ByVal strCRID As String, ByVal strERecord As String) As DataSet

            Dim strCommand As String = "usp_Upd_UEcr5"
            Dim strTableName As String = "UEcr5"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@ERecord", strERecord)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改資深巡輔員回饋
        Public Function usp_Upd_UEcr6(ByVal strCRID As String, ByVal strDSuggest As String) As DataSet

            Dim strCommand As String = "usp_Upd_UEcr6"
            Dim strTableName As String = "UEcr6"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@DSuggest", strDSuggest)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改輔導教授回饋
        Public Function usp_Upd_UEcr7(ByVal strCRID As String, ByVal strESuggest As String) As DataSet

            Dim strCommand As String = "usp_Upd_UEcr7"
            Dim strTableName As String = "UEcr7"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@ESuggest", strESuggest)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改教保輔導紀錄資料表狀態
        Public Function usp_Upd_UEcr8(ByVal strCRID As String, ByVal strState As String) As DataSet

            Dim strCommand As String = "usp_Upd_UEcr8"
            Dim strTableName As String = "UEcr8"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@State", strState)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改教保輔導紀錄資料表((條件：輔導紀錄編號))教保輔導重點、備註
        Public Function usp_Upd_UEcr9(ByVal strCRID As String, ByVal strEFocus As String, ByVal strERemark As String) As DataSet

            Dim strCommand As String = "usp_Upd_UEcr9"
            Dim strTableName As String = "UEcr9"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@EFocus", strEFocus)
            dbCheck.htParam.Add("@ERemark", strERemark)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改教保輔導紀錄資料表((條件：輔導紀錄編號))教保輔導重點、備註
        Public Function usp_Upd_UEcr10(ByVal strCRID As String, ByVal strCloseDate As String) As DataSet

            Dim strCommand As String = "usp_Upd_UEcr10"
            Dim strTableName As String = "UEcr10"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@CloseDate", strCloseDate)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改輔導紀錄草稿、初稿、完稿((條件：年、月、要修改狀態、原來狀態))草稿改初稿 @State1=1  @State1=0 。 初稿改完稿 @State1=2  @State1=1
        Public Function usp_Upd_UEcr11(ByVal strYear As String, ByVal strMonth As String, ByVal strState1 As String, ByVal strState2 As String) As DataSet

            Dim strCommand As String = "usp_Upd_UEcr11"
            Dim strTableName As String = "UEcr11"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@State1", strState1)
            dbCheck.htParam.Add("@State2", strState2)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''修改症狀、人數
        Public Function usp_Upd_UEch1(ByVal strCRID As String, ByVal strSymptom As String, ByVal strHealthNum As String) As DataSet
            Dim strCommand As String = "usp_Upd_UEch1"
            Dim strTableName As String = "UEch1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@Symptom", strSymptom)
            dbCheck.htParam.Add("@HealthNum", strHealthNum)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改內容、指標編號
        Public Function usp_Upd_UEcs1(ByVal strCRID As String, ByVal strCRSID As String, ByVal strSuggest As String, ByVal strTID As String) As DataSet
            Dim strCommand As String = "usp_Upd_UEcs1"
            Dim strTableName As String = "UEcs1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@CRSID", strCRSID)
            dbCheck.htParam.Add("@Suggest", strSuggest)
            dbCheck.htParam.Add("@TID", strTID)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改內容、指標編號、追蹤狀況編號
        Public Function usp_Upd_UEcs2(ByVal strCRID As String, ByVal strCRSID As String, ByVal strSuggest As String, ByVal strTID As String, ByVal strTSID As String) As DataSet
            Dim strCommand As String = "usp_Upd_UEcs2"
            Dim strTableName As String = "UEcs2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@CRSID", strCRSID)
            dbCheck.htParam.Add("@Suggest", strSuggest)
            dbCheck.htParam.Add("@TID", strTID)
            dbCheck.htParam.Add("@TSID", strTSID)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改追蹤狀況編號、結案日期、結案內容
        Public Function usp_Upd_UEcs3(ByVal strCRID As String, ByVal strCRSID As String, ByVal strNewTSID As String, ByVal strCloseDate As String, ByVal strCloseSuggest As String) As DataSet
            Dim strCommand As String = "usp_Upd_UEcs3"
            Dim strTableName As String = "UEcs3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@CRSID", strCRSID)
            dbCheck.htParam.Add("@NewTSID", strNewTSID)
            dbCheck.htParam.Add("@CloseDate", strCloseDate)
            dbCheck.htParam.Add("@CloseSuggest", strCloseSuggest)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改教保輔導優點與改善策略建議資料表((條件：輔導紀錄編號、建議編號))最新追蹤狀況編號、改善原因編號、結案日期、結案內容
        Public Function usp_Upd_UEcs4(ByVal strCRID As String, ByVal strCRSID As String, ByVal strNewTSID As String,
                                       ByVal strIRID As String, ByVal strCloseDate As String, ByVal strCloseSuggest As String) As DataSet
            Dim strCommand As String = "usp_Upd_UEcs4"
            Dim strTableName As String = "UEcs4"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@CRSID", strCRSID)
            dbCheck.htParam.Add("@NewTSID", strNewTSID)
            dbCheck.htParam.Add("@IRID", strIRID)
            dbCheck.htParam.Add("@CloseDate", strCloseDate)
            dbCheck.htParam.Add("@CloseSuggest", strCloseSuggest)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢教保輔導優點與改善策略建議資料表--建議追蹤缺點未結案((條件：帳號，學期時間起迄，SType =2 ，CloseDate='1911-01-01'))本學期追蹤缺點未結案，計算該學期EduCR_Suggest總比數
        Public Function usp_Sel_SEcs5(ByVal strA01 As String, ByVal strYeardate1 As String, ByVal strYeardate2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcs5"
            Dim strTableName As String = "SEcs5"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@Yeardate1", strYeardate1)
            dbCheck.htParam.Add("@Yeardate2", strYeardate2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢教保輔導優點與改善策略建議資料表--建議追蹤缺點已結案((條件：帳號，學期時間起迄，SType =2 ，CloseDate<>'1911-01-01'))本學期追蹤缺點未結案，計算該學期EduCR_Suggest總比數
        Public Function usp_Sel_SEcs6(ByVal strA01 As String, ByVal strYeardate1 As String, ByVal strYeardate2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcs6"
            Dim strTableName As String = "SEcs6"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@Yeardate1", strYeardate1)
            dbCheck.htParam.Add("@Yeardate2", strYeardate2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''查詢追蹤彙整表(依照舊版系統，非修復版本的格式)((條件：帳號，學期時間起迄，SType =2 ，EduCR_Member.GCIdentity='C' ))
        Public Function usp_Sel_SEcs7(ByVal strA01 As String, ByVal stryear1 As String, ByVal stryear2 As String) As DataSet
            Dim strCommand As String = "usp_Sel_SEcs7"
            Dim strTableName As String = "SEcs7"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@year1", stryear1)
            dbCheck.htParam.Add("@year2", stryear2)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function



        '''修改結案日期
        Public Function usp_Upd_UOld1(ByVal strCRID As String, ByVal strOldCRID As String, ByVal strOldCRSID As String, ByVal strCloseDate As String) As DataSet
            Dim strCommand As String = "usp_Upd_UOld1"
            Dim strTableName As String = "UOld1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@OldCRID", strOldCRID)
            dbCheck.htParam.Add("@OldCRSID", strOldCRSID)
            dbCheck.htParam.Add("@CloseDate", strCloseDate)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改指標名稱
        Public Function usp_Upd_UTgt1(ByVal strTVersion As String, ByVal strTID As String, ByVal strTName As String) As DataSet
            Dim strCommand As String = "usp_Upd_UTgt1"
            Dim strTableName As String = "UTgt1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TVersion", strTVersion)
            dbCheck.htParam.Add("@TID", strTID)
            dbCheck.htParam.Add("@TName", strTName)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改指標名稱
        Public Function usp_Upd_UTgt2() As DataSet
            Dim strCommand As String = "usp_Upd_UTgt2"
            Dim strTableName As String = "UTgt2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改所設定的版本(TYorN = 1)
        Public Function usp_Upd_UTgt3(ByVal strTVersion As String) As DataSet
            Dim strCommand As String = "usp_Upd_UTgt3"
            Dim strTableName As String = "UTgt3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TVersion", strTVersion)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改追蹤狀況名稱、追蹤狀況定義
        Public Function usp_Upd_UTst1(ByVal strTSVersion As String, ByVal strTSID As String, ByVal strTSName As String, ByVal strTSDef As String) As DataSet
            Dim strCommand As String = "usp_Upd_UTst1"
            Dim strTableName As String = "UTst1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TSVersion", strTSVersion)
            dbCheck.htParam.Add("@TSID", strTSID)
            dbCheck.htParam.Add("@TSName", strTSName)
            dbCheck.htParam.Add("@TSDef", strTSDef)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改((條件：版本、編號))改善原因名稱、改善原因定義
        Public Function usp_Upd_UIpr1(ByVal strIRVersion As String, ByVal strIRID As String, ByVal strIRName As String, ByVal strIRDef As String) As DataSet
            Dim strCommand As String = "usp_Upd_UIpr1"
            Dim strTableName As String = "UIpr1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@IRVersion", strIRVersion)
            dbCheck.htParam.Add("@IRID", strIRID)
            dbCheck.htParam.Add("@IRName", strIRName)
            dbCheck.htParam.Add("@IRDef", strIRDef)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改全部改善原因已經指定的版本變成沒採用(將IRYorN = 1 的變成 IRYorN = 0 )
        Public Function usp_Upd_UIpr2() As DataSet
            Dim strCommand As String = "usp_Upd_UIpr2"
            Dim strTableName As String = "UIpr2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''修改所設定的版本(IRYorN = 1)
        Public Function usp_Upd_UIpr3(ByVal strIRVersion As String) As DataSet
            Dim strCommand As String = "usp_Upd_UIpr3"
            Dim strTableName As String = "UIpr3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@IRVersion", strIRVersion)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改全部指標已經指定的版本變成沒採用(將TSYorN = 1 的變成 TSYorN = 0 )
        Public Function usp_Upd_UTst2() As DataSet
            Dim strCommand As String = "usp_Upd_UTst2"
            Dim strTableName As String = "UTst2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改所設定的版本(TSYorN = 1)
        Public Function usp_Upd_UTst3(ByVal strTSVersion As String) As DataSet
            Dim strCommand As String = "usp_Upd_UTst3"
            Dim strTableName As String = "UTst3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TSVersion", strTSVersion)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改園所名稱、是否有效、公私立、園所地址、園所電話、備註
        Public Function usp_Upd_UGtn1(ByVal strCityID As String, ByVal strGID As String, ByVal strGName As String, ByVal strGEff As String, ByVal strGPP As String, ByVal strGAddress As String, ByVal strGTEL As String, ByVal strGRemark As String) As DataSet
            Dim strCommand As String = "usp_Upd_UGtn1"
            Dim strTableName As String = "UGtn1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GName", strGName)
            dbCheck.htParam.Add("@GEff", strGEff)
            dbCheck.htParam.Add("@GPP", strGPP)
            dbCheck.htParam.Add("@GAddress", strGAddress)
            dbCheck.htParam.Add("@GTEL", strGTEL)
            dbCheck.htParam.Add("@GRemark", strGRemark)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改班級名稱、是否有效、班別、2~3歲人數、3~4歲人數、4~5歲人數、5歲人數
        Public Function usp_Upd_UGtc1(ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String, ByVal strGCName As String, ByVal strGCEff As String, ByVal strGCType As String, ByVal strGCPeople2 As String, ByVal strGCPeople3 As String, ByVal strGCPeople4 As String, ByVal strGCPeople5 As String) As DataSet
            Dim strCommand As String = "usp_Upd_UGtc1"
            Dim strTableName As String = "UGtc1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            dbCheck.htParam.Add("@GCName", strGCName)
            dbCheck.htParam.Add("@GCEff", strGCEff)
            dbCheck.htParam.Add("@GCType", strGCType)
            dbCheck.htParam.Add("@GCPeople2", strGCPeople2)
            dbCheck.htParam.Add("@GCPeople3", strGCPeople3)
            dbCheck.htParam.Add("@GCPeople4", strGCPeople4)
            dbCheck.htParam.Add("@GCPeople5", strGCPeople5)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''((條件：縣市、園所編號、班級編號))僅修改5歲人數(填報巡迴輔導紀錄用)
        Public Function usp_Upd_UGtc2(ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String, ByVal strGCPeople5 As String) As DataSet
            Dim strCommand As String = "usp_Upd_UGtc2"
            Dim strTableName As String = "UGtc2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            dbCheck.htParam.Add("@GCPeople5", strGCPeople5)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改系統公告資料表
        Public Function usp_Upd_UNtF1(ByVal strNFID As String, ByVal strTimes As String, ByVal strTopic As String, ByVal strContents As String) As DataSet
            Dim strCommand As String = "usp_Upd_UNtF1"
            Dim strTableName As String = "UNtF1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@NFID", strNFID)
            dbCheck.htParam.Add("@Times", strTimes)
            dbCheck.htParam.Add("@Topic", strTopic)
            dbCheck.htParam.Add("@Contents", strContents)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改問卷管理資料表
        Public Function usp_Upd_UQme1(ByVal strQID As String, ByVal strQName As String, ByVal strQContents As String, ByVal strQStarttime As String, ByVal strQEndtime As String) As DataSet
            Dim strCommand As String = "usp_Upd_UQme1"
            Dim strTableName As String = "UQme1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QName", strQName)
            dbCheck.htParam.Add("@QContents", strQContents)
            dbCheck.htParam.Add("@QStarttime", strQStarttime)
            dbCheck.htParam.Add("@QEndtime", strQEndtime)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改問卷題目分類
        Public Function usp_Upd_UQcf1(ByVal strQID As String, ByVal strQCID As String, ByVal strQCName As String, ByVal strQCcontents As String) As DataSet
            Dim strCommand As String = "usp_Upd_UQcf1"
            Dim strTableName As String = "UQcf1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            dbCheck.htParam.Add("@QCName", strQCName)
            dbCheck.htParam.Add("@QCcontents", strQCcontents)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改問卷題目
        Public Function usp_Upd_UQue1(ByVal strQID As String, ByVal strQCID As String, ByVal strQTID As String,
                                      ByVal strQTName As String, ByVal strQTtype As String, ByVal strQTchoice As String,
                                      ByVal strQTOther As String) As DataSet
            Dim strCommand As String = "usp_Upd_UQue1"
            Dim strTableName As String = "UQue1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            dbCheck.htParam.Add("@QTID", strQTID)
            dbCheck.htParam.Add("@QTName", strQTName)
            dbCheck.htParam.Add("@QTtype", strQTtype)
            dbCheck.htParam.Add("@QTchoice", strQTchoice)
            dbCheck.htParam.Add("@QTOther", strQTOther)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改問卷選項
        Public Function usp_Upd_UQop1(ByVal strQID As String, ByVal strQTID As String, ByVal strQOID As String,
                                      ByVal strQCID As String, ByVal strQOName As String, ByVal strQOjump As String) As DataSet
            Dim strCommand As String = "usp_Upd_UQop1"
            Dim strTableName As String = "UQop1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QTID", strQTID)
            dbCheck.htParam.Add("@QOID", strQOID)
            dbCheck.htParam.Add("@QCID", strQCID)
            dbCheck.htParam.Add("@QOName", strQOName)
            dbCheck.htParam.Add("@QOjump", strQOjump)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''((條件:問卷編號、帳號))更新最後填答時間、填答狀態
        Public Function usp_Upd_UQur1(ByVal strQTime As String, ByVal strQState As String, ByVal strQID As String,
                                      ByVal strA01 As String, ByVal strIPAddress As String) As DataSet
            Dim strCommand As String = "usp_Upd_UQur1"
            Dim strTableName As String = "UQur1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QTime", strQTime)
            dbCheck.htParam.Add("@QState", strQState)
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@IPAddress", strIPAddress)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改每月非教保輔導資料表((條件：縣市編號，報告日期、項目編號))日期、開始時間、結束時間、機構名稱、處理業務簡述、備註、巡輔員帳號
        Public Function usp_Upd_URnecrm1(ByVal strCityID As String, ByVal strReportDate As String, ByVal strReportID As String,
                                      ByVal strDay As String, ByVal strStartTime As String, ByVal strEndTime As String,
                                         ByVal strInstitution As String, ByVal strBusiness As String, ByVal strReportRemark As String,
                                         strA01 As String) As DataSet
            Dim strCommand As String = "usp_Upd_URnecrm1"
            Dim strTableName As String = "URnecrm1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@ReportID", strReportID)
            dbCheck.htParam.Add("@Day", strDay)
            dbCheck.htParam.Add("@StartTime", strStartTime)
            dbCheck.htParam.Add("@EndTime", strEndTime)
            dbCheck.htParam.Add("@Institution", strInstitution)
            dbCheck.htParam.Add("@Business", strBusiness)
            dbCheck.htParam.Add("@ReportRemark", strReportRemark)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改輔導教授每月填報園所資料表((條件：縣市編號，報告日期、項目編號、輔導教授帳號))日期、園所編號、備註內容(訪視心得/省思)
        Public Function usp_Upd_UReportPG1(ByVal strCityID As String, ByVal strReportDate As String, ByVal strReportID As String,
                                      ByVal strDay As String, ByVal strGID As String, ByVal strPContent As String, ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Upd_UReportPG1"
            Dim strTableName As String = "UReportPG1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@ReportID", strReportID)
            dbCheck.htParam.Add("@Day", strDay)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@PContent", strPContent)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改每月上傳檔案((條件：西元年int、月份int、縣市、檔案編號))類型、檔案名稱、巡輔員帳號
        Public Function usp_Upd_UUlf1(ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String,
                                      ByVal strFileID As String, ByVal strType As String, ByVal strFileName As String,
                                      ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Upd_UUlf1"
            Dim strTableName As String = "UUlf1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@FileID", strFileID)
            dbCheck.htParam.Add("@Type", strType)
            dbCheck.htParam.Add("@FileName", strFileName)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改每月會議記錄出席教授資料表((條件：西元年int、月份int、縣市、編號))帳號(出席教授)、日期、會議名稱、巡輔員帳號
        Public Function usp_Upd_UAmt1(ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String,
                                      ByVal strAMID As String, ByVal strA01 As String, ByVal strDay As String,
                                      ByVal strMeetName As String, ByVal strAA01 As String) As DataSet
            Dim strCommand As String = "usp_Upd_UAmt1"
            Dim strTableName As String = "UAmt1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@AMID", strAMID)
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@Day", strDay)
            dbCheck.htParam.Add("@MeetName", strMeetName)
            dbCheck.htParam.Add("@AA01", strAA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改帳號與身分證字號轉換
        Public Function usp_Upd_UUUU(ByVal strA011 As String, ByVal strA012 As String) As DataSet
            Dim strCommand As String = "usp_Upd_UUUU"
            Dim strTableName As String = "UUUU"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A011", strA011)
            dbCheck.htParam.Add("@A012", strA012)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改園所代碼ACT
        Public Function usp_Upd_UUUU2(ByVal strA01 As String, ByVal strA12 As String) As DataSet
            Dim strCommand As String = "usp_Upd_UUUU2"
            Dim strTableName As String = "UUUU2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@A12", strA12)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改行政輔導紀錄資料表((條件：回報日期、前次輔導紀錄編號、前次建議編號、縣市、幼兒園、班級))追蹤狀況編號、追蹤方式、內容
        Public Function usp_Upd_UAmr1(ByVal strReportDate As String, ByVal strOldCRID As String, ByVal strOldCRSID As String,
                                      ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String, ByVal strTSID As String,
                                      ByVal strTrackWay As String, ByVal strContents As String) As DataSet
            Dim strCommand As String = "usp_Upd_UAmr1"
            Dim strTableName As String = "UAmr1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@OldCRID", strOldCRID)
            dbCheck.htParam.Add("@OldCRSID", strOldCRSID)
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            dbCheck.htParam.Add("@TSID", strTSID)
            dbCheck.htParam.Add("@TrackWay", strTrackWay)
            dbCheck.htParam.Add("@Contents", strContents)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''修改生日
        Public Function usp_Upd_UUUU3(ByVal strA01 As String, ByVal strA06 As String) As DataSet
            Dim strCommand As String = "usp_Upd_UUUU3"
            Dim strTableName As String = "UUUU3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            dbCheck.htParam.Add("@A06", strA06)

            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function



        '''刪除教保輔導人員資料表
        Public Function usp_Del_DAcc1(ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Del_DAcc1"
            Dim strTableName As String = "DAcc1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除教保輔導人員資料表
        Public Function usp_Del_DEcm(ByVal strCRID As String) As DataSet
            Dim strCommand As String = "usp_Del_DEcm"
            Dim strTableName As String = "DEcm"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除教保輔導優點與改善策略建議資料表
        Public Function usp_Del_DEch1(ByVal strCRID As String, ByVal strSymptom As String) As DataSet
            Dim strCommand As String = "usp_Del_DEch1"
            Dim strTableName As String = "DEch1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@Symptom", strSymptom)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除教保輔導優點與改善策略建議資料表
        Public Function usp_Del_DEcs1(ByVal strCRID As String, ByVal strCRSID As String) As DataSet
            Dim strCommand As String = "usp_Del_DEcs1"
            Dim strTableName As String = "DEcs1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            dbCheck.htParam.Add("@CRSID", strCRSID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除教保輔導前次建議追蹤資料表
        Public Function usp_Del_DOld(ByVal strCRID As String) As DataSet
            Dim strCommand As String = "usp_Del_DOld"
            Dim strTableName As String = "DOld"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除教保輔導紀錄資料表
        Public Function usp_Del_DEcr1(ByVal strCRID As String) As DataSet
            Dim strCommand As String = "usp_Del_DEcr1"
            Dim strTableName As String = "DEcr1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除教保輔導紀錄資料表
        Public Function usp_Del_DEch2(ByVal strCRID As String) As DataSet
            Dim strCommand As String = "usp_Del_DEch2"
            Dim strTableName As String = "DEch2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除教保輔導優點與改善策略建議資料表
        Public Function usp_Del_DEcs2(ByVal strCRID As String) As DataSet
            Dim strCommand As String = "usp_Del_DEcs2"
            Dim strTableName As String = "DEcs2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CRID", strCRID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除指標資料
        Public Function usp_Del_DTgt1(ByVal strTVersion As String, ByVal strTID As String) As DataSet
            Dim strCommand As String = "usp_Del_DTgt1"
            Dim strTableName As String = "DTgt1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TVersion", strTVersion)
            dbCheck.htParam.Add("@TID", strTID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除指標資料
        Public Function usp_Del_DTst1(ByVal strTSVersion As String, ByVal strTSID As String) As DataSet
            Dim strCommand As String = "usp_Del_DTst1"
            Dim strTableName As String = "DTst1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@TSVersion", strTSVersion)
            dbCheck.htParam.Add("@TSID", strTSID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''追蹤改善原因資料表
        Public Function usp_Del_DIpr1(ByVal strIRVersion As String, ByVal strIRID As String) As DataSet
            Dim strCommand As String = "usp_Del_DIpr1"
            Dim strTableName As String = "DIpr1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@IRVersion", strIRVersion)
            dbCheck.htParam.Add("@IRID", strIRID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除園所班級資料表
        Public Function usp_Del_DGtc1(ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String) As DataSet
            Dim strCommand As String = "usp_Del_DGtc1"
            Dim strTableName As String = "DGtc1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除園所班級資料表((條件：縣市、園所編號))
        Public Function usp_Del_DGtc2(ByVal strCityID As String, ByVal strGID As String) As DataSet
            Dim strCommand As String = "usp_Del_DGtc2"
            Dim strTableName As String = "DGtc2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除園所班級人員資料表
        Public Function usp_Del_DGcm1(ByVal strCityID As String, ByVal strGID As String, ByVal strGCID As String) As DataSet
            Dim strCommand As String = "usp_Del_DGcm1"
            Dim strTableName As String = "DGcm1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除園所班級人員資料表(條件：縣市、園所)
        Public Function usp_Del_DGcm2(ByVal strCityID As String, ByVal strGID As String) As DataSet
            Dim strCommand As String = "usp_Del_DGcm2"
            Dim strTableName As String = "DGcm2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除園所資料表
        Public Function usp_Del_DGtn1(ByVal strCityID As String, ByVal strGID As String) As DataSet
            Dim strCommand As String = "usp_Del_DGtn1"
            Dim strTableName As String = "DGtn1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''刪除系統公告資料表
        Public Function usp_Del_DNtf1(ByVal strNFID As String) As DataSet
            Dim strCommand As String = "usp_Del_DNtf1"
            Dim strTableName As String = "DNtf1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@NFID", strNFID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除問卷管理資料表
        Public Function usp_Del_DQme1(ByVal strQID As String) As DataSet
            Dim strCommand As String = "usp_Del_DQme1"
            Dim strTableName As String = "DQme1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function


        '''刪除問卷題目分類
        Public Function usp_Del_DQcf1(ByVal strQID As String, ByVal strQCID As String) As DataSet
            Dim strCommand As String = "usp_Del_DQcf1"
            Dim strTableName As String = "DQcf1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除問卷題目分類(依照問卷編號)
        Public Function usp_Del_DQcf2(ByVal strQID As String) As DataSet
            Dim strCommand As String = "usp_Del_DQcf2"
            Dim strTableName As String = "DQcf2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除選擇的問卷選項
        Public Function usp_Del_DQop1(ByVal strQID As String, ByVal strQTID As String, ByVal strQOID As String) As DataSet
            Dim strCommand As String = "usp_Del_DQop1"
            Dim strTableName As String = "DQop1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QTID", strQTID)
            dbCheck.htParam.Add("@QOID", strQOID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除選擇的問卷選項
        Public Function usp_Del_DQop2(ByVal strQID As String, ByVal strQCID As String, ByVal strQTID As String) As DataSet
            Dim strCommand As String = "usp_Del_DQop2"
            Dim strTableName As String = "DQop2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            dbCheck.htParam.Add("@QTID", strQTID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除選擇的問卷選項
        Public Function usp_Del_DQop3(ByVal strQID As String) As DataSet
            Dim strCommand As String = "usp_Del_DQop3"
            Dim strTableName As String = "DQop3"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除題目
        Public Function usp_Del_DQue1(ByVal strQID As String, ByVal strQCID As String, ByVal strQTID As String) As DataSet
            Dim strCommand As String = "usp_Del_DQue1"
            Dim strTableName As String = "DQue1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            dbCheck.htParam.Add("@QTID", strQTID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除題目(條件：問卷編號)
        Public Function usp_Del_DQue2(ByVal strQID As String) As DataSet
            Dim strCommand As String = "usp_Del_DQue2"
            Dim strTableName As String = "DQue2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''新增問卷填寫紀錄(問卷編號、題目分類編號、題目編號、選項編號、帳號、其他文字)
        Public Function usp_Del_DQrc1(ByVal strQID As String, ByVal strQCID As String, ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Del_DQrc1"
            Dim strTableName As String = "DQrc1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@QCID", strQCID)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除問卷填寫紀錄((條件：問卷編號))
        Public Function usp_Del_DQrc2(ByVal strQID As String) As DataSet
            Dim strCommand As String = "usp_Del_DQrc2"
            Dim strTableName As String = "DQrc2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除問卷受訪人員((條件：問卷編號))
        Public Function usp_Del_DQur1(ByVal strQID As String) As DataSet
            Dim strCommand As String = "usp_Del_DQur1"
            Dim strTableName As String = "DQur1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除問卷受訪人員((條件：問卷編號、帳號))
        Public Function usp_Del_DQur2(ByVal strQID As String, ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Del_DQur2"
            Dim strTableName As String = "DQur2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@QID", strQID)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除月輔導報告資料表((條件：縣市編號、報告日期、巡輔員帳號))
        Public Function usp_Del_DRecrm1(ByVal strCityID As String, ByVal strReportDate As String, ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Del_DRecrm1"
            Dim strTableName As String = "DRecrm1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除每月非教保報告資料表((條件：縣市編號、報告日期、項目編號))
        Public Function usp_Del_DRnecrm1(ByVal strCityID As String, ByVal strReportDate As String, ByVal strReportID As String) As DataSet
            Dim strCommand As String = "usp_Del_DRnecrm1"
            Dim strTableName As String = "DRnecrm1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@ReportID", strReportID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除每月上傳檔案資料表
        Public Function usp_Del_DUlf1(ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String, ByVal strFileID As String) As DataSet
            Dim strCommand As String = "usp_Del_DUlf1"
            Dim strTableName As String = "DUlf1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@FileID", strFileID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除每月上傳檔案資料表(該年月份的上傳檔案資訊)
        Public Function usp_Del_DUlf2(ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Del_DUlf2"
            Dim strTableName As String = "DUlf2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除每月會議記錄出席教授資料表
        Public Function usp_Del_DAmt1(ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String, ByVal strAMID As String) As DataSet
            Dim strCommand As String = "usp_Del_DAmt1"
            Dim strTableName As String = "DAmt1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@AMID", strAMID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除每月會議記錄出席教授資料表
        Public Function usp_Del_DAmt2(ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Del_DAmt2"
            Dim strTableName As String = "DAmt2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除月輔導報告資料表((條件：縣市編號、報告日期、輔導教授帳號))
        Public Function usp_Del_DReportPA1(ByVal strCityID As String, ByVal strReportDate As String, ByVal strA01 As String) As DataSet
            Dim strCommand As String = "usp_Del_DReportPA1"
            Dim strTableName As String = "DReportPA1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@A01", strA01)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除輔導教授每月填報園所資料表((條件：縣市編號、報告日期、項目編號))
        Public Function usp_Del_DReportPG1(ByVal strCityID As String, ByVal strReportDate As String, ByVal strReportID As String) As DataSet
            Dim strCommand As String = "usp_Del_DReportPG1"
            Dim strTableName As String = "DReportPG1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@ReportID", strReportID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除輔導教授每月填報園所資料表((條件：縣市編號、報告日期))
        Public Function usp_Del_DReportPG2(ByVal strYear As String, ByVal strMonth As String, ByVal strCityID As String) As DataSet
            Dim strCommand As String = "usp_Del_DReportPG2"
            Dim strTableName As String = "DReportPG2"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@Year", strYear)
            dbCheck.htParam.Add("@Month", strMonth)
            dbCheck.htParam.Add("@CityID", strCityID)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function

        '''刪除行政輔導紀錄資料表((條件：回報日期、縣市、園所編號、班級編號、追蹤方式))
        Public Function usp_Del_DAmr1(ByVal strReportDate As String, ByVal strCityID As String, ByVal strGID As String,
                                      ByVal strGCID As String, ByVal strTrackWay As String) As DataSet
            Dim strCommand As String = "usp_Del_DAmr1"
            Dim strTableName As String = "DAmr1"
            Dim dbCheck As New PISCP.DB.ClassDB
            Dim dsDone As New DataSet()
            Dim spParams As SqlParameter() = Nothing

            dbCheck.htParam.Clear()
            dbCheck.htParam.Add("@ReportDate", strReportDate)
            dbCheck.htParam.Add("@CityID", strCityID)
            dbCheck.htParam.Add("@GID", strGID)
            dbCheck.htParam.Add("@GCID", strGCID)
            dbCheck.htParam.Add("@TrackWay", strTrackWay)
            spParams = dbCheck.getSpParams(0)
            dsDone = dbCheck.RunProcedure(strCommand, spParams, strTableName)

            Return dsDone
        End Function




    End Class
End Namespace
