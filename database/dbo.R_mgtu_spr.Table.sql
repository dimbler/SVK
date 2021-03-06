USE [ecdep]
GO
/****** Object:  Справочник для работы скрипта шифрования ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[R_mgtu_spr](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[dt_z] [datetime] NULL CONSTRAINT [DF_R_mgtu_spr_dt_z]  DEFAULT (getdate()),
	[name_oi] [varchar](100) NULL CONSTRAINT [DF_R_mgtu_spr_name_oi]  DEFAULT ([dbo].[sp_get_namepol]()),
	[suser] [varchar](100) NULL CONSTRAINT [DF_R_mgtu_spr_suser]  DEFAULT ((suser_name()+' : ')+host_name()),
	[file_mask] [varchar](500) NULL,
	[name_otch] [varchar](100) NULL, /** Наименование отчета **/
	[mail] [varchar](150) NULL, /** Адрес отправки **/
	[kvit] [varchar](max) NULL, /** Алгоритм создания маршрутной квитанции **/
	[max_files] [int] NULL, /** Максимальное количество файлов в архиве **/
	[max_size] [int] NULL, /** Максимальный размер файла без разбивки **/
	[encr_id] [nchar](4) NULL, /** Номер ключа шифрования **/
	[abonent] [varchar](150) NULL, /** Наименование абонента **/
	[encr_path] [varchar](150) NULL, /** Путь к ключу шифрования **/
	[sign_path] [varchar](150) NULL, /** Путь к ключу подписи **/
	[archiv_encr_id] [nchar](4) NULL, /** Номер ключа шифрования архива **/
	[archiv_encr_path] [varchar](150) NULL, /** Путь к ключу шифрования архива **/
	[file_ukep] [varchar](150) NULL, /** Итентификатор ключа унифицированной электронной подписи **/
	[archiv_ukep] [varchar](150) NULL, /** Итентификатор ключа унифицированной электронной подписи архива **/
	[archive_name] [varchar](150) NULL, /** Итентификатор или название файла архива **/
	[file_flow] [varchar](255) NULL, /** Необходимые операции выполняемые в скрипте **/
	[archiv_flow] [varchar](255) NULL, /** Необходимые операции с архивом выполняемые в скрипте **/
	[prn] [varchar](50) NULL, /** Путь к принтеру для печати результатов операции **/
	[pandion] [varchar](200) NULL /** Адреса для уведомления сотрудников **/
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
SET IDENTITY_INSERT [dbo].[R_mgtu_spr] ON 

INSERT [dbo].[R_mgtu_spr] ([id], [dt_z], [name_oi], [suser], [file_mask], [name_otch], [mail], [kvit], [max_files], [max_size], [encr_id], [abonent], [encr_path], [sign_path], [archiv_encr_id], [archiv_encr_path], [file_ukep], [archiv_ukep], [archive_name], [file_flow], [archiv_flow], [Extensions], [prn], [def_path], [pandion]) VALUES (34, CAST(N'2016-12-26 12:24:57.067' AS DateTime), NULL, N'владелец записи', NULL, N'024', N'crypt@ext-gate.svk.mskgtu.cbr.ru', NULL, 1, 5, N'0200', N'ТУ БР', N'путь к ключу шифрования', N'путь к ключу подписи', NULL, NULL, NULL, NULL, NULL, N'KA;ENCR', NULL, NULL, N'принтер для печати отчета', NULL, NULL)
SET IDENTITY_INSERT [dbo].[R_mgtu_spr] OFF
