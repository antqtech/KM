-- =============================================
-- Author:	Carol
-- Create date: 2023/08/01
-- Description:	查詢選手資料
-- =============================================
CREATE Procedure usp_Sel_SPlayer1

	(
          @PID varchar(50)
	)

AS
SELECT PID, PName_Zh, PName_En from Player WHERE PID = @PID or PPhone = @PID
GO
/*
GRANT EXEC ON Stored_Procedure_Name TO PUBLIC

GO
*/