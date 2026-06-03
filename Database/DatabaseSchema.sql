-- ComServ Africa POS System - MS Access Database Schema
-- This SQL can be used as reference for database structure

-- Table: tblUsers
CREATE TABLE tblUsers (
    UserID AUTOINCREMENT PRIMARY KEY,
    Username TEXT NOT NULL UNIQUE,
    Password TEXT NOT NULL,
    Role TEXT NOT NULL CHECK (Role IN ('Admin', 'Teller')),
    CreatedDate DATETIME DEFAULT NOW(),
    IsActive BOOLEAN DEFAULT TRUE
);

-- Table: tblProducts
CREATE TABLE tblProducts (
    ProductID AUTOINCREMENT PRIMARY KEY,
    ProductName TEXT NOT NULL,
    Specs TEXT,
    UnitPrice DECIMAL(10,2) NOT NULL,
    StockLevel INTEGER NOT NULL DEFAULT 0,
    Warranty TEXT DEFAULT '6 Months',
    CreatedDate DATETIME DEFAULT NOW(),
    LastModified DATETIME DEFAULT NOW()
);

-- Table: tblSales
CREATE TABLE tblSales (
    SaleID AUTOINCREMENT PRIMARY KEY,
    TellerName TEXT NOT NULL,
    SaleDate DATETIME DEFAULT NOW(),
    SubtotalAmount DECIMAL(10,2) NOT NULL,
    TaxAmount DECIMAL(10,2) NOT NULL,
    TotalAmount DECIMAL(10,2) NOT NULL,
    PaymentMethod TEXT,
    CustomerID INTEGER
);

-- Table: tblSaleItems
CREATE TABLE tblSaleItems (
    ItemID AUTOINCREMENT PRIMARY KEY,
    SaleID INTEGER NOT NULL,
    ProductID INTEGER NOT NULL,
    Quantity INTEGER NOT NULL,
    PriceAtSale DECIMAL(10,2) NOT NULL,
    LineTotal DECIMAL(10,2) NOT NULL
);

-- Table: tblCustomers
CREATE TABLE tblCustomers (
    CustomerID AUTOINCREMENT PRIMARY KEY,
    CustomerName TEXT NOT NULL,
    ContactNumber TEXT,
    EmailAddress TEXT,
    Address TEXT,
    CreatedDate DATETIME DEFAULT NOW()
);

-- Table: tblPromotions
CREATE TABLE tblPromotions (
    PromotionID AUTOINCREMENT PRIMARY KEY,
    PromotionName TEXT NOT NULL,
    DiscountPercentage DECIMAL(5,2) NOT NULL,
    StartDate DATETIME NOT NULL,
    EndDate DATETIME NOT NULL,
    ApplicableProducts TEXT,
    MinimumPurchase DECIMAL(10,2),
    IsActive BOOLEAN DEFAULT TRUE,
    CreatedDate DATETIME DEFAULT NOW()
);

-- Table: tblBlacklist
CREATE TABLE tblBlacklist (
    BlacklistID AUTOINCREMENT PRIMARY KEY,
    CustomerID INTEGER NOT NULL,
    Reason TEXT NOT NULL,
    DateAdded DATETIME DEFAULT NOW(),
    AddedBy TEXT NOT NULL
);

-- Table: tblWhitelist
CREATE TABLE tblWhitelist (
    WhitelistID AUTOINCREMENT PRIMARY KEY,
    CustomerID INTEGER NOT NULL,
    DiscountPercentage DECIMAL(5,2) DEFAULT 0,
    ValidFrom DATETIME NOT NULL,
    ValidTo DATETIME,
    CreatedDate DATETIME DEFAULT NOW()
);

-- Table: tblReturns
CREATE TABLE tblReturns (
    ReturnID AUTOINCREMENT PRIMARY KEY,
    SaleID INTEGER NOT NULL,
    ItemID INTEGER NOT NULL,
    ReturnQuantity INTEGER NOT NULL,
    ReturnReason TEXT NOT NULL,
    ReturnAmount DECIMAL(10,2) NOT NULL,
    ReturnDate DATETIME DEFAULT NOW(),
    ProcessedBy TEXT NOT NULL,
    Approved BOOLEAN DEFAULT FALSE
);

-- Table: tblQuotations
CREATE TABLE tblQuotations (
    QuotationID AUTOINCREMENT PRIMARY KEY,
    CustomerID INTEGER,
    QuotationDate DATETIME DEFAULT NOW(),
    ValidUntil DATETIME NOT NULL,
    SubtotalAmount DECIMAL(10,2) NOT NULL,
    TaxAmount DECIMAL(10,2) NOT NULL,
    TotalAmount DECIMAL(10,2) NOT NULL,
    CreatedBy TEXT NOT NULL,
    Status TEXT DEFAULT 'Pending'
);

-- Table: tblQuotationItems
CREATE TABLE tblQuotationItems (
    QuotationItemID AUTOINCREMENT PRIMARY KEY,
    QuotationID INTEGER NOT NULL,
    ProductID INTEGER NOT NULL,
    Quantity INTEGER NOT NULL,
    PriceAtQuotation DECIMAL(10,2) NOT NULL,
    LineTotal DECIMAL(10,2) NOT NULL
);