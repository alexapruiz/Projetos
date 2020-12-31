USE [VENDAS]
GO

/****** Object:  Table [dbo].[PROCESSADOR]    Script Date: 25/08/2020 19:45:53 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PROCESSADOR]') AND type in (N'U'))
DROP TABLE [dbo].[PROCESSADOR]
GO

/****** Object:  Table [dbo].[PROCESSADOR]    Script Date: 25/08/2020 19:45:53 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[PROCESSADOR]
(
	[ID_PROCESSADOR] [int] NOT NULL,
	[ID_SOCKET] [int] NOT NULL,
	[ID_TIPO_MEMORIA] [int] NOT NULL,
	[ID_MARCA] [int] NOT NULL,
	[DESC_PROCESSADOR] [nchar](30) NOT NULL,
	[VALOR_NACIONAL] [money] NOT NULL,
	[FORNECEDOR_NACIONAL] [nchar](20) NOT NULL,
	[URL_NACIONAL] [char](999) NOT NULL,
	[VALOR_IMPORTADO] [money] NOT NULL,
	[FORNECEDOR_IMPORTADO] [nchar](20) NOT NULL,
	[URL_IMPORTADO] [char](999) NOT NULL
)

INSERT INTO [dbo].[PROCESSADOR] ([ID_PROCESSADOR] ,[ID_SOCKET] ,[ID_TIPO_MEMORIA],[ID_MARCA],[DESC_PROCESSADOR],[VALOR_NACIONAL],[FORNECEDOR_NACIONAL],[URL_NACIONAL],[VALOR_IMPORTADO],[FORNECEDOR_IMPORTADO],[URL_IMPORTADO] ) 
			VALUES
           (1,2,9,1,'Core i3 9100F',
		   489.98,'Pichau','https://www.pichau.com.br/hardware/processador-intel-core-i3-9100-quad-core-3-6ghz-4-2ghz-turbo-6mb-cache-lga1151-bx80684i39100f',
		   444.27,'Ali Express','https://pt.aliexpress.com/item/4001226821708.html?spm=a2g0o.productlist.0.0.1c5645dc1xJ4ur&algo_pvid=954104d3-ae32-4105-9a31-8df40f2e8ee7&algo_expid=954104d3-ae32-4105-9a31-8df40f2e8ee7-0&btsid=0ab6fa7b15983972285825582e5669&ws_ab_test=searchweb0_0,searchweb201602_,searchweb201603_')
GO
