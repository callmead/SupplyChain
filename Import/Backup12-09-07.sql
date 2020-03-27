-- MySQL dump 10.9
--
-- Host: localhost    Database: ak_inv
-- ------------------------------------------------------
-- Server version	4.1.13a-nt

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `customer`
--

DROP TABLE IF EXISTS `customer`;
CREATE TABLE `customer` (
  `Customer_ID` varchar(20) NOT NULL default '',
  `Date` date default NULL,
  `Name` varchar(30) default NULL,
  `CNIC_No` varchar(15) default NULL,
  `Address` varchar(50) default NULL,
  `Occupation` varchar(30) default NULL,
  `Phone_No` varchar(15) default NULL,
  `Mobile_No` varchar(15) default NULL,
  `Other_No` varchar(15) default NULL,
  `Total_Bills_Amount` int(11) default NULL,
  `Total_Due` int(11) default NULL,
  `Remarks` varchar(50) default NULL,
  PRIMARY KEY  (`Customer_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `customer`
--


/*!40000 ALTER TABLE `customer` DISABLE KEYS */;
LOCK TABLES `customer` WRITE;
INSERT INTO `customer` VALUES ('C2007829135114','2007-08-29','sasasa','-','-','-','-','-','-',0,0,'-');
UNLOCK TABLES;
/*!40000 ALTER TABLE `customer` ENABLE KEYS */;

--
-- Table structure for table `customer_account`
--

DROP TABLE IF EXISTS `customer_account`;
CREATE TABLE `customer_account` (
  `TID` varchar(20) NOT NULL default '',
  `Customer_ID` varchar(20) default NULL,
  `Date` date default NULL,
  `Invoice_No` varchar(20) default NULL,
  `Total_Amount` int(11) default NULL,
  `Payment_Mode` varchar(15) default NULL,
  `Amount_Paid` int(11) default NULL,
  `Amount_Due` int(11) default NULL,
  `Remarks` varchar(50) default NULL,
  PRIMARY KEY  (`TID`),
  KEY `fk_Inv1_No` (`Invoice_No`),
  CONSTRAINT `fk_Inv1_No` FOREIGN KEY (`Invoice_No`) REFERENCES `invoice` (`Invoice_No`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `customer_account`
--


/*!40000 ALTER TABLE `customer_account` DISABLE KEYS */;
LOCK TABLES `customer_account` WRITE;
UNLOCK TABLES;
/*!40000 ALTER TABLE `customer_account` ENABLE KEYS */;

--
-- Table structure for table `invoice`
--

DROP TABLE IF EXISTS `invoice`;
CREATE TABLE `invoice` (
  `TID` varchar(20) NOT NULL default '',
  `Invoice_No` varchar(20) default NULL,
  `Product_ID` varchar(20) default NULL,
  `Quantity` int(11) default NULL,
  `Price` int(11) default NULL,
  `Net_Total` int(11) default NULL,
  PRIMARY KEY  (`TID`),
  KEY `fk_Inv_No` (`Invoice_No`),
  KEY `fk_prod_id` (`Product_ID`),
  CONSTRAINT `fk_Inv_No` FOREIGN KEY (`Invoice_No`) REFERENCES `sales` (`Invoice_No`),
  CONSTRAINT `fk_prod_id` FOREIGN KEY (`Product_ID`) REFERENCES `stock` (`Product_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `invoice`
--


/*!40000 ALTER TABLE `invoice` DISABLE KEYS */;
LOCK TABLES `invoice` WRITE;
UNLOCK TABLES;
/*!40000 ALTER TABLE `invoice` ENABLE KEYS */;

--
-- Table structure for table `login`
--

DROP TABLE IF EXISTS `login`;
CREATE TABLE `login` (
  `User` varchar(15) default NULL,
  `Password` varchar(10) default NULL,
  `Account_Type` varchar(15) default NULL,
  `Name` varchar(20) default NULL,
  `Designation` varchar(20) default NULL,
  `Remarks` varchar(50) default NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `login`
--


/*!40000 ALTER TABLE `login` DISABLE KEYS */;
LOCK TABLES `login` WRITE;
INSERT INTO `login` VALUES ('admin','admin','Admin','Admin User','Administration','-'),('manager','manager','Manager','Manager User','Management','-'),('salesman','salesman','Salesman','Sales User','Sales Dept','-'),('Shahid','1234','Salesman','Shahid Jamil','Salesman','-');
UNLOCK TABLES;
/*!40000 ALTER TABLE `login` ENABLE KEYS */;

--
-- Table structure for table `po_details`
--

DROP TABLE IF EXISTS `po_details`;
CREATE TABLE `po_details` (
  `TID` varchar(20) NOT NULL default '',
  `PO_No` varchar(20) default NULL,
  `Product` varchar(20) default NULL,
  `Product_Type` varchar(20) default NULL,
  `Product_Size` varchar(10) default NULL,
  `Quantity` int(11) default NULL,
  `Description` varchar(30) default NULL,
  PRIMARY KEY  (`TID`),
  KEY `fk_PO1_No` (`PO_No`),
  CONSTRAINT `fk_PO1_No` FOREIGN KEY (`PO_No`) REFERENCES `purchase_order` (`PO_No`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `po_details`
--


/*!40000 ALTER TABLE `po_details` DISABLE KEYS */;
LOCK TABLES `po_details` WRITE;
UNLOCK TABLES;
/*!40000 ALTER TABLE `po_details` ENABLE KEYS */;

--
-- Table structure for table `purchase_order`
--

DROP TABLE IF EXISTS `purchase_order`;
CREATE TABLE `purchase_order` (
  `PO_No` varchar(20) NOT NULL default '',
  `Date` date default NULL,
  `Supplier_ID` varchar(20) default NULL,
  `Delivery_Date` date default NULL,
  `Remarks` varchar(50) default NULL,
  PRIMARY KEY  (`PO_No`),
  KEY `fk_supp_id1` (`Supplier_ID`),
  CONSTRAINT `fk_supp_id1` FOREIGN KEY (`Supplier_ID`) REFERENCES `supplier` (`Supplier_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `purchase_order`
--


/*!40000 ALTER TABLE `purchase_order` DISABLE KEYS */;
LOCK TABLES `purchase_order` WRITE;
UNLOCK TABLES;
/*!40000 ALTER TABLE `purchase_order` ENABLE KEYS */;

--
-- Table structure for table `receivings`
--

DROP TABLE IF EXISTS `receivings`;
CREATE TABLE `receivings` (
  `TID` varchar(20) NOT NULL default '',
  `Date` date default NULL,
  `PO_No` varchar(20) default NULL,
  `Product_ID` varchar(20) default NULL,
  `Quantity` int(11) default NULL,
  `Price` int(11) default NULL,
  `Price_per_unit` int(11) default NULL,
  `Remarks` varchar(50) default NULL,
  PRIMARY KEY  (`TID`),
  KEY `fk_PO_No1` (`PO_No`),
  KEY `fk_prod_id3` (`Product_ID`),
  CONSTRAINT `fk_PO_No1` FOREIGN KEY (`PO_No`) REFERENCES `purchase_order` (`PO_No`),
  CONSTRAINT `fk_prod_id3` FOREIGN KEY (`Product_ID`) REFERENCES `stock` (`Product_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `receivings`
--


/*!40000 ALTER TABLE `receivings` DISABLE KEYS */;
LOCK TABLES `receivings` WRITE;
UNLOCK TABLES;
/*!40000 ALTER TABLE `receivings` ENABLE KEYS */;

--
-- Table structure for table `sales`
--

DROP TABLE IF EXISTS `sales`;
CREATE TABLE `sales` (
  `Invoice_No` varchar(20) NOT NULL default '',
  `Date` date default NULL,
  `Salesman` varchar(20) default NULL,
  `Customer_ID` varchar(20) default NULL,
  `Grand_Total` int(11) default NULL,
  `Discount` int(11) default NULL,
  `Payment_Mode` varchar(15) default NULL,
  `Amount_Paid` int(11) default NULL,
  `Amount_Due` int(11) default NULL,
  `Remarks` varchar(50) default NULL,
  PRIMARY KEY  (`Invoice_No`),
  KEY `fk_cust_id` (`Customer_ID`),
  CONSTRAINT `fk_cust_id` FOREIGN KEY (`Customer_ID`) REFERENCES `customer` (`Customer_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `sales`
--


/*!40000 ALTER TABLE `sales` DISABLE KEYS */;
LOCK TABLES `sales` WRITE;
UNLOCK TABLES;
/*!40000 ALTER TABLE `sales` ENABLE KEYS */;

--
-- Table structure for table `stock`
--

DROP TABLE IF EXISTS `stock`;
CREATE TABLE `stock` (
  `Product_ID` varchar(20) NOT NULL default '',
  `Date` date default NULL,
  `Product` varchar(30) default NULL,
  `Product_Type` varchar(20) default NULL,
  `Product_Size` varchar(10) default NULL,
  `Stock_In_Hand` int(11) default NULL,
  `Description` varchar(25) default NULL,
  `Price_per_unit` int(11) default NULL,
  `ReOrder_Level` int(11) default NULL,
  `Remarks` varchar(50) default NULL,
  PRIMARY KEY  (`Product_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `stock`
--


/*!40000 ALTER TABLE `stock` DISABLE KEYS */;
LOCK TABLES `stock` WRITE;
INSERT INTO `stock` VALUES ('P2007829151035','2007-08-29','Diyar 404','Lamination','8x4',85,'Premier',1050,25,'-'),('P2007829151213','2007-08-29','Masa 222','Lamination','8x4',75,'Premier',1050,15,'-'),('P2007829151533','2007-08-29','Classic Mix','Lamination','-',125,'Premier',1125,10,'-'),('P2007829151745','2007-08-29','Diyar 233','Lamination','-',18,'PPI',1040,10,'-'),('P2007829151940','2007-08-29','Masa 350','Lamination','-',22,'PPI',1040,15,'-'),('P2007829152043','2007-08-29','Diyar 350','lamination','-',12,'PPI',1040,15,'-'),('P2007829152357','2007-08-29','Brown','Winboard','8x4',88,'Alpha wood',860,15,'-'),('P200782915252','2007-08-29','Masawa','Winboard','8x4',5,'Alpha Wood',1080,15,'-'),('P2007829152550','2007-08-29','Simble Winboard','Winboard','8x4',17,'-',775,10,'-'),('P2007829152653','2007-08-29','Lasani 2.5','Lasani','8x4',60,'Imported',313,10,'-'),('P2007829152953','2007-08-29','Lasani 5.5 mm','Lasani','8x4',20,'Imported',500,5,'-'),('P2007829153041','2007-08-29','Lasani 8 mm','Lasani','8x4',15,'Imported',800,5,'-'),('P2007829153229','2007-08-29','Lasani 11 mm','Lasani','8x4',3,'Imported',980,5,'-'),('P2007829153451','2007-08-29','Lasani 15 mm','Lasani','8x4',3,'Imported',1350,5,'-'),('P2007829153626','2007-08-29','Lasani 16 mm','Lasani','8x4',30,'A Grade',1200,10,'-'),('P20078291554','2007-08-29','Diyar 233','Lamination','8 x 4',45,'Premier Formica Peshawer',1050,25,'-'),('P200782915727','2007-08-29','Diyar 350','Lamination','8x4',6,'Premier Formica Peshawer',1050,25,'-');
UNLOCK TABLES;
/*!40000 ALTER TABLE `stock` ENABLE KEYS */;

--
-- Table structure for table `supplier`
--

DROP TABLE IF EXISTS `supplier`;
CREATE TABLE `supplier` (
  `Supplier_ID` varchar(20) NOT NULL default '',
  `Date` date default NULL,
  `Name` varchar(30) default NULL,
  `Company` varchar(20) default NULL,
  `Contact_Person` varchar(30) default NULL,
  `Address` varchar(50) default NULL,
  `Office_No` varchar(15) default NULL,
  `Mobile_No` varchar(15) default NULL,
  `Other_No` varchar(15) default NULL,
  `Fax_No` varchar(15) default NULL,
  `Total_Bills_Amount` int(11) default NULL,
  `Total_Due` int(11) default NULL,
  `Remarks` varchar(50) default NULL,
  PRIMARY KEY  (`Supplier_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `supplier`
--


/*!40000 ALTER TABLE `supplier` DISABLE KEYS */;
LOCK TABLES `supplier` WRITE;
INSERT INTO `supplier` VALUES ('S200782919256','2007-08-29','Premier','Premier Formica','Naeem ur Rehman','Premier Formica Company Peshawer','--','03018917716','03458587562','-',286100,0,'-'),('S200782919834','2007-08-29','Peshawer Partical','Peshawer Partical In','Fazel-e-Rabbi','Peshawer Partical Industry Hayatabad Peshawer Paki','--','03008598337','-','-',283140,0,'-');
UNLOCK TABLES;
/*!40000 ALTER TABLE `supplier` ENABLE KEYS */;

--
-- Table structure for table `supplier_account`
--

DROP TABLE IF EXISTS `supplier_account`;
CREATE TABLE `supplier_account` (
  `TID` varchar(20) NOT NULL default '',
  `Supplier_ID` varchar(20) default NULL,
  `Date` date default NULL,
  `PO_No` varchar(20) default NULL,
  `Total_Amount` int(11) default NULL,
  `Payment_Mode` varchar(15) default NULL,
  `Paid_Amount` int(11) default NULL,
  `Due_Amount` int(11) default NULL,
  `Remarks` varchar(50) default NULL,
  PRIMARY KEY  (`TID`),
  KEY `fk_supp_id2` (`Supplier_ID`),
  KEY `fk_PO_No` (`PO_No`),
  CONSTRAINT `fk_PO_No` FOREIGN KEY (`PO_No`) REFERENCES `purchase_order` (`PO_No`),
  CONSTRAINT `fk_supp_id2` FOREIGN KEY (`Supplier_ID`) REFERENCES `supplier` (`Supplier_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `supplier_account`
--


/*!40000 ALTER TABLE `supplier_account` DISABLE KEYS */;
LOCK TABLES `supplier_account` WRITE;
UNLOCK TABLES;
/*!40000 ALTER TABLE `supplier_account` ENABLE KEYS */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

