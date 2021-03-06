/****** Таблица ошибок в работе скрипта требующих уведомления оператора ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[R_mgtu_err](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[err_message] [varchar](200) NULL,
	[err_flag] [int] NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
SET IDENTITY_INSERT [dbo].[R_mgtu_err] ON 

INSERT [dbo].[R_mgtu_err] ([id], [err_message], [err_flag]) VALUES (1, N'ОЭС не принят в обработку', 102)
INSERT [dbo].[R_mgtu_err] ([id], [err_message], [err_flag]) VALUES (2, N'дата версии программы должна быть равна', 101)
INSERT [dbo].[R_mgtu_err] ([id], [err_message], [err_flag]) VALUES (3, N'Результат контроля: 1 - протокол содержит сообщения о предупреждениях', 1)
INSERT [dbo].[R_mgtu_err] ([id], [err_message], [err_flag]) VALUES (5, N'ERRORS_ES', 103)
INSERT [dbo].[R_mgtu_err] ([id], [err_message], [err_flag]) VALUES (6, N'не принят', 105)
INSERT [dbo].[R_mgtu_err] ([id], [err_message], [err_flag]) VALUES (7, N'Не удалось', 106)
INSERT [dbo].[R_mgtu_err] ([id], [err_message], [err_flag]) VALUES (8, N'Данный отчет заменил отчет с датой регистрации ', 201)
SET IDENTITY_INSERT [dbo].[R_mgtu_err] OFF
