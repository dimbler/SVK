/****** Object:  Техническая таблица связывающая шифрованные сообщения с файлами отправленными на шифрование ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[R_mgtu_ref](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[task_guid] [varchar](100) NULL,
	[file_guid] [varchar](100) NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO

