Public Function Accept(oVideo As VideoDetail, TypeID As Long, EmployeeID As Long, AppID As Long, EditingLogID As Long, Tags As List(Of Long), ExtraMessage As String) As ActionResponse
        Dim oAR As New ActionResponse
        oAR.Action = "Accept"
        Dim oLog As New CPShared.Log(oAR.Action)

        Dim o As VideoVersionDetail = CBO.Video.Version.GetDetail(oVideo.VideoID, TypeID)
        If o Is Nothing Then
            oAR.Message = "Video does not exits"
            oLog.AddProblem(oAR.Message)
            LogError(oVideo.VideoID, TypeID, EmployeeID, CPS.Logs.GetEditingLog(oLog), AppID)
            oAR.MessageCustomer = "An error occurred."
            oAR.IsValid = False
            Return oAR
        End If
        oLog.Add("Type", o.VersionName)

        Dim oDT As New DataTable
        oDT.Load(CDL.Video.Version.ByVideoSelect(oVideo.VideoID))
        If oDT Is Nothing Then
            oAR.Message = "Error occured. Try again"
            oLog.Add("Problem", "Could not get all versions of the video")
            LogError(oVideo.VideoID, TypeID, EmployeeID, CPS.Logs.GetEditingLog(oLog), AppID)
            oAR.MessageCustomer = "An error occurred."
            oAR.IsValid = False
            Return oAR
        End If

        Dim oTags As TagLists = GetTags(Tags, oVideo.VideoID)
        For Each t As Tag In oTags.Deleted
            oLog.Add("Tag deleted", t.TagName)
        Next

        For Each t As Tag In oTags.Inserted
            oLog.Add("Tag inserted", t.TagName)
        Next

        If ExtraMessage <> "" Then
            oLog.Add("Message", ExtraMessage)
        End If

        Using oC As SqlConnection = New SqlConnection(CDL.DB.Cnn)
            Dim oT As SqlTransaction
            oC.Open()
            oT = oC.BeginTransaction
            Try
                CDL.Video.Update(oVideo.VideoID, oVideo.FilmID, oVideo.TypeID, oVideo.SourceID, oVideo.FileSize, oVideo.EmployeeID, oVideo.StageID, oVideo.FormatName, oVideo.IsLive, oVideo.FormatLongName, oVideo.Duration, oVideo.StreamCount, oVideo.Bitrate, oVideo.ProgramCount, oVideo.FileName, oVideo.SourceLink, oVideo.FrameSize, oVideo.LanguageTypeID, oT:=oT, oC:=oC)
                CDL.Video.Version.StageUpdate(oVideo.VideoID, TypeID, enVideoVersionStage.ToUpload, oC:=oC, oT:=oT)

                If o.IsMain = True Then
                    For Each dr As DataRow In oDT.Rows
                        If TypeID <> dr("TypeID") Then
                            CDL.Video.Version.Update(oVideo.VideoID, dr("TypeID"), 2, o.MachineID, o.AudioIndex, o.VideoIndex, o.FPS, o.Deinterlace, o.Crop, o.AspectRatio, o.StartTime, o.EndTime, o.EncodeTypeID, False, False, o.AviSynthAudioIndex, o.AviSynthVideoIndex, o.FinalWidth, o.FinalHeight, o.FromFile, oC:=oC, oT:=oT)
                        End If
                    Next
                End If


                For Each t As Tag In oTags.Deleted
                    CDL.Video.Tag.Delete(oVideo.VideoID, t.TagID, oT:=oT, oC:=oC)
                Next

                For Each t As Tag In oTags.Inserted
                    CDL.Video.Tag.Insert(oVideo.VideoID, t.TagID, EmployeeID, oT:=oT, oC:=oC)
                Next

                CDL.Video.AcceptanceLog.FinishedUpdate(EditingLogID, oC:=oC, oT:=oT)
                CDL.Video.Log.Insert(oVideo.VideoID, EmployeeID, CPS.Logs.GetEditingLog(oLog), oT:=oT, oC:=oC)
                oT.Commit()
            Catch ex As Exception
                oT.Rollback()
                LogError(oVideo.VideoID, TypeID, EmployeeID, CPS.Logs.GetEditingLog(oLog), AppID)
                oAR.Message = ex.Message
                oAR.MessageCustomer = "An error occurred."
                oAR.IsValid = False
                Return oAR
            Finally
                If oC.State <> ConnectionState.Closed Then oC.Close()
            End Try
        End Using
        oAR.IsValid = True
        Return oAR
    End Function
