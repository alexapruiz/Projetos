/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE */;
/*!40101 SET SQL_MODE='STRICT_TRANS_TABLES,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES */;
/*!40103 SET SQL_NOTES='ON' */;


CREATE DATABASE `planeta` /*!40100 DEFAULT CHARACTER SET latin1 */;
USE `planeta`;
CREATE TABLE `clientes` (
  `Codigo` int(11) NOT NULL AUTO_INCREMENT,
  `Nome` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`Codigo`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
CREATE TABLE `item` (
  `codigo` int(11) NOT NULL AUTO_INCREMENT,
  `tipo` int(11) NOT NULL DEFAULT '0',
  `descricao` varchar(255) DEFAULT NULL,
  `valor_unit` decimal(10,2) DEFAULT NULL,
  PRIMARY KEY (`codigo`)
) ENGINE=InnoDB AUTO_INCREMENT=35 DEFAULT CHARSET=latin1;
CREATE TABLE `item_pedido` (
  `codigo_pedido` int(11) NOT NULL,
  `codigo_item` int(11) NOT NULL DEFAULT '0',
  `qtde` int(11) NOT NULL DEFAULT '0',
  `tema` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`codigo_pedido`,`codigo_item`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
CREATE TABLE `pedido` (
  `codigo` int(11) NOT NULL AUTO_INCREMENT,
  `data_entrega` varchar(10) NOT NULL DEFAULT '',
  `hora_entrega` varchar(5) NOT NULL DEFAULT '00:00',
  `cliente` int(11) NOT NULL DEFAULT '0',
  PRIMARY KEY (`codigo`)
) ENGINE=InnoDB AUTO_INCREMENT=38 DEFAULT CHARSET=latin1;
CREATE TABLE `tipo_item` (
  `Codigo` int(11) NOT NULL AUTO_INCREMENT,
  `Descricao` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`Codigo`)
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=latin1;

CREATE DEFINER=`root`@`localhost` PROCEDURE `selecionapedidos`(`Codigo` int(11))
BEGIN
select * from pedido;
END;


/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;
/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
