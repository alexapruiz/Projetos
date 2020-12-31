USE [VENDAS]
GO

/****** Object:  Table [dbo].[PLACA_MAE]    Script Date: 25/08/2020 19:46:12 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PLACA_MAE]') AND type in (N'U'))
DROP TABLE [dbo].[PLACA_MAE]
GO

/****** Object:  Table [dbo].[PLACA_MAE]    Script Date: 25/08/2020 19:46:12 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[PLACA_MAE]
(
	[ID_PLACA_MAE] [int] NOT NULL,
	[ID_SOCKET] [int] NOT NULL,
	[MODELO_PLACA] [nchar](50) NOT NULL,
	[CHIPSET_PLACA] [nchar](10) NOT NULL,
	[VALOR_NACIONAL] [money] NOT NULL,
	[FORNECEDOR_NACIONAL] [nchar](20) NOT NULL,
	[URL_NACIONAL] [char](999) NOT NULL,
	[VALOR_IMPORTADO] [money] NOT NULL,
	[FORNECEDOR_IMPORTADO] [nchar](20) NOT NULL,
	[URL_IMPORTADO] [char](999) NOT NULL

)
GO


USE [VENDAS]
GO

INSERT INTO [dbo].[PLACA_MAE]
           (
			[ID_PLACA_MAE],
			[ID_SOCKET],
			[MODELO_PLACA],
			[CHIPSET_PLACA],
			[VALOR_NACIONAL],
			[FORNECEDOR_NACIONAL],
			[URL_NACIONAL],
			[VALOR_IMPORTADO],
			[FORNECEDOR_IMPORTADO],
			[URL_IMPORTADO]
			)
     VALUES
           (1,2,'ASUS TUF H310M-PLUS GAMING/BR','H310',
		   599.02,'Pichau','https://www.pichau.com.br/hardware/placa-mae-asus-prime-h310m-e-r2-0-br-ddr4-socket-lga1151-chipset-intel-h310',
		   534.70,'Ali Express','https://pt.aliexpress.com/item/32982791507.html?spm=a2g0o.productlist.0.0.5c0d213cuidUjD&algo_pvid=3b4007fd-b58d-4f9f-a287-8659fc73a76f&algo_expid=3b4007fd-b58d-4f9f-a287-8659fc73a76f-0&btsid=0ab6d69515983987671672589e98a2&ws_ab_test=searchweb0_0,searchweb201602_,searchweb201603_'
		   )