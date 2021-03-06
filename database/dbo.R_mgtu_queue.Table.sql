/****** Object:  Таблица очереди на шифрование сообщений ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[R_mgtu_queue](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[queue_id] [varchar](50) NULL,
	[sended_filename] [varchar](255) NULL,
	[status] [varchar](20) NULL,
	[processed_time] [datetime] NULL,
	[sended_time] [datetime] NULL,
	[received_time_k1] [datetime] NULL,
	[received_path_k1] [varchar](255) NULL,
	[received_time_k2] [datetime] NULL,
	[received_path_k2] [varchar](255) NULL,
	[file_path] [varchar](255) NULL,
	[received_file_k1] [varchar](255) NULL,
	[received_file_k2] [varchar](255) NULL,
	[status_k1] [varchar](255) NULL,
	[status_k2] [varchar](255) NULL,
	[name_otch] [varchar](100) NULL,
	[err_flag_k1] [int] NULL,
	[err_flag_k2] [int] NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
