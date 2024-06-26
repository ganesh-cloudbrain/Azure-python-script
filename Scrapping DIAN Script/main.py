from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from playwright.sync_api import sync_playwright
from email.utils import parsedate_tz, mktime_tz
from Rules import RulesAccountNumbers,CODIGO
from datetime import datetime,timezone
from bs4 import BeautifulSoup
import requests
import pandas as pd
import zipfile
import shutil
import base64
import time
import html
import os
import re

Parties             = []
Documents           = []
AccountingDocs      = []
Customers           =[{
                        "CustomerID":"",	
                        "Name":"",	
                        "DV":"",	
                        "City":"",	
                        "Country":"",	
                        "Mail":"",	
                        "Phone":"",	
                        "Address":"",	
                        "TaxRegime":"",	
                        "SpecialRegime":"",	
                        "PrincipalActiviy":"",	
                        "TaxPayerAttribute":"",	
                        "ElectronicSupplier":"",	
                        "Token":"",	
                        "TestID":"",	
                        "Decimals":"",	
                        "Software":"",	
                        "Token":"",

}]

def checkEmpty(text):
    if text == "" or text is None:
        return 0.0
    else:
        return text

def standardizeDicts(listOfDicts):
    allKeys = set(key for d in listOfDicts for key in d.keys())

    for d in listOfDicts:
        missingKeys = allKeys - d.keys()
        for key in missingKeys:
            d[key] = ""

    return listOfDicts


def writeDictstoExcel(listOfDicts, fileName):
    
    df0 = pd.DataFrame(Customers)
    df1 = pd.DataFrame(Parties)
    df2 = pd.DataFrame(Documents)
    df3 = pd.DataFrame(AccountingDocs)
    df4 = pd.DataFrame(RulesAccountNumbers)

    # Create a Pandas Excel writer using openpyxl as the engine
    writer = pd.ExcelWriter(fileName, engine='openpyxl')

    # Write each DataFrame to a different worksheet
    df0.to_excel(writer, sheet_name='Customers', index=False)
    df1.to_excel(writer, sheet_name='Parties', index=False)
    df2.to_excel(writer, sheet_name='Documents', index=False)
    df3.to_excel(writer, sheet_name='AccountingDoc', index=False)
    df4.to_excel(writer, sheet_name='RulesAccountNumber', index=False)

    writer.close()

def getIDType(IDType):

    IDTypeDict = {
        "Registro civil"                         : "10910092",
        "Tarjeta de identidad"                   : "10910093",
        "Cédula de ciudadanía"                   : "10910094",
        "Tarjeta de extranjería"                 : "10910095",
        "Cédula de extranjería"                  : "10910096",
        "NIT"                                    : "10910097",
        "Pasaporte"                              : "10910098",
        "Tipo de documento desconocido"          : "10910366",
        "Documento de identificación extranjero" : "10910394",
        "Nit de otro país"                       : "10910402",
        "NIUP"                                   : "10910403",
        "PEP"                                    : "10930954",
        "PPT"                                    : "10930955"
    }

    try:
        IdentificationType=IDTypeDict[IDType]
        return IdentificationType
    except KeyError:
        print(f"ID Type : {IDType} not found !!!\nUsing Default ID Type : Cédula de ciudadanía")
        IdentificationType=IDTypeDict["Cédula de ciudadanía"]
        return IdentificationType

def calculateDV(party):
    
    try:
        companyIDstr = f"{int(party['CompanyID']):015d}"
        
        weights = [71, 67, 59, 53, 47, 43, 41, 37, 29, 23, 19, 17, 13, 7, 3]
        
        weightedSum = sum(int(companyIDstr[i]) * weights[i] for i in range(15))
        
        residual = weightedSum % 11
        
        if residual == 0:
            DV = 0
        elif residual == 1:
            DV = 1
        else:
            DV = 11 - residual
        
        return DV
    except (TypeError,ValueError) :
        return ""


def findInvoiceLineTags(elements):
    invoiceTags = []
    
    for xmlString in elements:
        element = BeautifulSoup(xmlString,"xml")
        invoice = element.findAll("Invoice")
        invoiceTags.extend(invoice)

    return invoiceTags


def extractCData(filePath):
    with open(filePath, 'r') as file:
        content = file.read()
        cdataRegex = re.compile(r'<!\[CDATA\[(.*?)\]\]>', re.DOTALL)
        matches = cdataRegex.findall(content)
        if len(matches) < 1:
            return[content]
        
        return matches

def getAccountNumbersfromTable(RulesAccountNumbers,customerID,companyID,Type,Description,Percent):
    for accountNumber in RulesAccountNumbers:
        if str(Type).strip() == "Expenses":
            if  str(accountNumber["CompanyID"]).strip() == str(customerID).strip() and str(Type).strip() == str(accountNumber["Type"]).strip() and str(accountNumber["ExpensesKeyWord"]).strip() == "":
                return accountNumber
            elif  str(accountNumber["CompanyID"]).strip() == str(customerID).strip() and str(Type).strip() == str(accountNumber["Type"]).strip() and str(accountNumber["ExpensesKeyWord"]).strip().lower() in str(Description).strip().lower(): 
                return accountNumber
        
        elif str(Type).strip() == "Tax":
            if  str(accountNumber["CompanyID"]).strip() == str(customerID).strip() and str(Type).strip() == str(accountNumber["Type"]).strip() and str(accountNumber["ExpensesKeyWord"]).strip().lower() in str(Description).strip().lower(): 
                return accountNumber
        
        elif str(Type).strip() == "WithHolding":
            # print(str(accountNumber["CompanyID"]).strip() == str(customerID).strip())
            # print(str(Type).strip() == str(accountNumber["Type"]).strip())
            # print(str(Type).strip())
            # print(str(accountNumber["Type"]).strip())
            # print(str(accountNumber["ExpensesKeyWord"]).strip().lower() in str(Description).strip().lower())
            # print(str(accountNumber["ExpensesKeyWord"]).strip().lower())
            # print(str(Description).strip().lower())
            
            # input()
            if str(accountNumber["CompanyID"]).strip() == str(customerID).strip() and str(Type).strip() == str(accountNumber["Type"]).strip() and str(accountNumber["ExpensesKeyWord"]).strip().lower() in str(Description).strip().lower():
                # print("WithHolding")
                # input()
                if type(accountNumber["Percent"]) == str: 
                    return accountNumber
                elif float(accountNumber["Percent"]) == float(Percent):
                    return accountNumber
    else:
        None
        



def parsingXMLFiles(XML_FILES_DIRECTORY_PATH,CCODE):

    for file in sorted(os.listdir(XML_FILES_DIRECTORY_PATH)):

        invoiceTags = findInvoiceLineTags(extractCData(os.path.join(XML_FILES_DIRECTORY_PATH,file)))
        for invoice in invoiceTags:
                """
                    AccountingSupplierParty
                """
                try:
                    ASP_registrationName                = invoice.select_one("AccountingSupplierParty>Party>PartyTaxScheme>RegistrationName").text
                except AttributeError:
                    ASP_registrationName                = ""

                try:
                    ASP_TaxLevelCode                    = invoice.select_one("AccountingSupplierParty>Party>PartyTaxScheme>TaxLevelCode").text
                except AttributeError:
                    ASP_TaxLevelCode                    = ""
                
                try:
                    ASP_CityID                          = invoice.select_one("AccountingSupplierParty>Party>PartyTaxScheme>RegistrationAddress>ID").text
                except AttributeError:
                    ASP_CityID                          = ""
                
                try:
                    ASP_CityName                        = invoice.select_one("AccountingSupplierParty>Party>PartyTaxScheme>RegistrationAddress>CityName").text
                except AttributeError:
                    ASP_CityName                        = ""
                
                try:
                    ASP_CountrySubentity                = invoice.select_one("AccountingSupplierParty>Party>PartyTaxScheme>RegistrationAddress>CountrySubentity").text
                except AttributeError:
                    ASP_CountrySubentity                = ""

                try:
                    ASP_CountrySubentityCode            = invoice.select_one("AccountingSupplierParty>Party>PartyTaxScheme>RegistrationAddress>CountrySubentityCode").text
                except AttributeError:
                    ASP_CountrySubentityCode            = ""
                
                try:
                    ASP_AddressLine                     = invoice.select_one("AccountingSupplierParty>Party>PartyTaxScheme>RegistrationAddress>AddressLine>Line").text
                except AttributeError:
                    ASP_AddressLine                     = ""

                try:
                    ASP_CountryIdentificationCode       = invoice.select_one("AccountingSupplierParty>Party>PartyTaxScheme>RegistrationAddress>Country>IdentificationCode").text
                except AttributeError:
                    ASP_CountryIdentificationCode       = ""
                
                try:
                    ASP_CountryName                     = invoice.select_one("AccountingSupplierParty>Party>PartyTaxScheme>RegistrationAddress>Country>Name").text
                except AttributeError:
                    ASP_CountryName                     = ""
                
                try:
                    ASP_Telephone                       = invoice.select_one("AccountingSupplierParty>Party>Contact>Telephone").text
                except AttributeError:
                    ASP_Telephone                       = ""
                
                try:
                    ASP_ElectronicMail                  = invoice.select_one("AccountingSupplierParty>Party>Contact>ElectronicMail").text
                except AttributeError:
                    ASP_ElectronicMail                  = ""

                try:
                    ASP_supplierID                      = invoice.select_one("AccountingSupplierParty>Party>PartyTaxScheme>CompanyID").text
                except AttributeError:
                    ASP_supplierID                      =  ""
                


                """
                AccountingCustomerParty
                """

                try:
                    ACP_registrationName                = invoice.select_one("AccountingCustomerParty>Party>PartyTaxScheme>RegistrationName").text
                except AttributeError:
                    ACP_registrationName                = ""

                try:
                    ACP_TaxLevelCode                    = invoice.select_one("AccountingCustomerParty>Party>PartyTaxScheme>TaxLevelCode").text
                except AttributeError:
                    ACP_TaxLevelCode                    = ""
                
                try:
                    ACP_CityID                          = invoice.select_one("AccountingCustomerParty>Party>PartyTaxScheme>RegistrationAddress>ID").text
                except AttributeError:
                    ACP_CityID                          = ""
                
                try:
                    ACP_CityName                        = invoice.select_one("AccountingCustomerParty>Party>PartyTaxScheme>RegistrationAddress>CityName").text
                except AttributeError:
                    ACP_CityName                        = ""
                
                try:
                    ACP_CountrySubentity                = invoice.select_one("AccountingCustomerParty>Party>PartyTaxScheme>RegistrationAddress>CountrySubentity").text
                except AttributeError:
                    ACP_CountrySubentity                = ""

                try:
                    ACP_CountrySubentityCode            = invoice.select_one("AccountingCustomerParty>Party>PartyTaxScheme>RegistrationAddress>CountrySubentityCode").text
                except AttributeError:
                    ACP_CountrySubentityCode            = ""
                
                try:
                    ACP_AddressLine                     = invoice.select_one("AccountingCustomerParty>Party>PartyTaxScheme>RegistrationAddress>AddressLine>Line").text
                except AttributeError:
                    ACP_AddressLine                     = ""

                try:
                    ACP_CountryIdentificationCode       = invoice.select_one("AccountingCustomerParty>Party>PartyTaxScheme>RegistrationAddress>Country>IdentificationCode").text
                except AttributeError:
                    ACP_CountryIdentificationCode       = ""
                
                try:
                    ACP_CountryName                     = invoice.select_one("AccountingCustomerParty>Party>PartyTaxScheme>RegistrationAddress>Country>Name").text
                except AttributeError:
                    ACP_CountryName                     = ""
                
                try:
                    ACP_Telephone                       = invoice.select_one("AccountingCustomerParty>Party>Contact>Telephone").text
                except AttributeError:
                    ACP_Telephone                       = ""
                
                try:
                    ACP_ElectronicMail                  = invoice.select_one("AccountingCustomerParty>Party>Contact>ElectronicMail").text
                except AttributeError:
                    ACP_ElectronicMail                  = ""

                try:
                    ACP_supplierID                      = invoice.select_one("AccountingCustomerParty>Party>PartyTaxScheme>CompanyID").text
                except AttributeError:
                    ACP_supplierID                      =  ""
                
                
                """
                    Global
                """

                try:
                    issueDate                       = invoice.select_one("IssueDate").text
                except AttributeError:
                    issueDate                       = ""
                
                try:
                    PaymentDueDate                  = invoice.select_one("PaymentMeans>PaymentDueDate").text
                except AttributeError:
                    PaymentDueDate                  = ""
                
                try:
                    PaymentMeansCode                = invoice.select_one("PaymentMeans>ID").text
                except AttributeError:
                    PaymentMeansCode                = ""

                try:
                    CalculationRate                 = invoice.select_one("PaymentMeans>CalculationRate").text
                except AttributeError:
                    CalculationRate                 = ""
                
                try:
                    OrderReference                  = invoice.select_one("OrderReference>ID").text
                except AttributeError:
                    OrderReference                  = "" 

                try:
                    Prefix                          = invoice.select_one("UBLExtensions>UBLExtension>ExtensionContent>DianExtensions>InvoiceControl>AuthorizedInvoices>Prefix").text
                except AttributeError:
                    Prefix                          = ""                               

                try:
                    documentID                      = invoice.select_one("ID").text
                except AttributeError:
                    documentID                      = ""
                
                try:
                    UUID                            = invoice.select_one("UUID").text
                except AttributeError:
                    UUID                            = ""
                

                ASP_Party= {
                    "CompanyID"   : ASP_supplierID.upper(),
                    "DV"          : "",
                    "Name"        : ASP_registrationName.upper(),	
                    "TaxLevel"    : ASP_TaxLevelCode.upper(),	
                    "IdCity"      : ASP_CityID.upper(),	
                    "CityName"    : ASP_CityName.upper(),	
                    "CityName2"   : ASP_CountrySubentity.upper(),	
                    "CityCode"    : ASP_CountrySubentityCode.upper(),	
                    "Address"     : ASP_AddressLine.upper(),	
                    "CountryCode" : ASP_CountryIdentificationCode.upper(),	
                    "CountryName" : ASP_CountryName.upper(),
                    "Phone"       : ASP_Telephone.upper(),
                    "Mail"        : ASP_ElectronicMail.upper(),
                }

                ASP_Party["DV"] = calculateDV(ASP_Party)
                
                if ASP_Party not in Parties:
                    Parties.append(ASP_Party)
                
                ACP_Party= {
                    "CompanyID"   : ACP_supplierID.upper(),
                    "DV"          : "",
                    "Name"        : ACP_registrationName.upper(),	
                    "TaxLevel"    : ACP_TaxLevelCode.upper(),	
                    "IdCity"      : ACP_CityID.upper(),	
                    "CityName"    : ACP_CityName.upper(),	
                    "CityName2"   : ACP_CountrySubentity.upper(),	
                    "CityCode"    : ACP_CountrySubentityCode.upper(),	
                    "Address"     : ACP_AddressLine.upper(),	
                    "CountryCode" : ACP_CountryIdentificationCode.upper(),	
                    "CountryName" : ACP_CountryName.upper(),
                    "Phone"       : ACP_Telephone.upper(),
                    "Mail"        : ACP_ElectronicMail.upper(),
                }
                
                ACP_Party["DV"] = calculateDV(ACP_Party)

                if ACP_Party not in Parties:
                    Parties.append(ACP_Party)

                if CCODE == ACP_supplierID:
                        CompanyID = ASP_supplierID
                        TYPE      = "Expenses"
                        
                else:
                    CompanyID = ACP_supplierID
                    TYPE      = "Income"
                
                
                
                WithHoldingTotalTax             = 0
                InvoiceLineTotalTax             = 0
                InvoiceLineExtensionTotalAmount = 0
                for invoiceLine in invoice.findAll("InvoiceLine"):

                    try:
                        invoiceLineItemDescription          = invoiceLine.select_one("Item>Description").text                
                    except AttributeError:
                        invoiceLineItemDescription          = ""
                    
                    try:
                        invoiceLineID                       = invoiceLine.select_one("ID").text                
                    except AttributeError:
                        invoiceLineID                       = ""
                    
                    try:
                        invoiceLineExtensionAmount          = invoiceLine.select_one("LineExtensionAmount").text                
                    except AttributeError:
                        invoiceLineExtensionAmount          = ""
                    
                    try:
                        invoiceLineTaxName                  = invoiceLine.select_one("TaxTotal>TaxSubtotal>TaxCategory>TaxScheme>Name").text                
                    except AttributeError:
                        invoiceLineTaxName                  = ""
                    
                    try:
                        invoiceLineTaxID                    = invoiceLine.select_one("TaxTotal>TaxSubtotal>TaxCategory>TaxScheme>ID").text                
                    except AttributeError:
                        invoiceLineTaxID                    = ""
                    
                    try:
                        invoiceLinecurrencyID               = invoiceLine.select_one("LineExtensionAmount")["currencyID"]                
                    except AttributeError:
                        invoiceLinecurrencyID               = ""
                    
                    try:
                        invoiceLinePercentage               = invoiceLine.select_one("TaxTotal>TaxSubtotal>TaxCategory>Percent").text
                    except AttributeError:
                        invoiceLinePercentage               = ""
                    
                    try:
                        invoiceLineTaxAmount                = invoiceLine.select_one("TaxTotal>TaxSubtotal>TaxAmount").text
                    except AttributeError:
                        invoiceLineTaxAmount                = ""
                    
                    try:
                        invoiceLineTaxableAmount            = invoiceLine.select_one("TaxTotal>TaxSubtotal>TaxableAmount").text
                    except AttributeError:
                        invoiceLineTaxableAmount            = ""

                    
                    AccountNumbers = getAccountNumbersfromTable(RulesAccountNumbers,CompanyID,CCODE,"Expenses",invoiceLineItemDescription,"")
                    if AccountNumbers is not None:
                        ExpensesAccount     = AccountNumbers["Account"] 
                    else:
                        ExpensesAccount     = "" 
                        
                    
                    
                    
                    Expenses= {
                        "Type"          : TYPE,
                        "DocType"       : "",	
                        "Number"        : documentID,	
                        "Date"          : issueDate,	
                        "AccountNumber" : ExpensesAccount,	
                        "CompanyID"     : CompanyID,	
                        "Prefix"        : Prefix,	
                        "SupplierDoc"   : documentID.replace(Prefix,""),	
                        "DueDate"       : "",	
                        "Description"   : invoiceLineItemDescription,	
                        "CostNumer"     : "",	
                        "Value"         : round(float(invoiceLineExtensionAmount),2)       if invoiceLineExtensionAmount       != "" else invoiceLineExtensionAmount,	
                        "Nature"        : "D",	
                        "TaxableAmount" : "",

                    }
                    AccountingDocs.append(Expenses)    
                    
                    
                    AccountNumbers = getAccountNumbersfromTable(RulesAccountNumbers,CompanyID,CCODE,"Tax",invoiceLineItemDescription,"")
                    if AccountNumbers is not None: 
                        AccountTax          = AccountNumbers["Account"]
                    else: 
                        AccountTax          = ""


                    TaxCalculation = (float(checkEmpty(invoiceLinePercentage)) / 100) * float(checkEmpty(invoiceLineTaxableAmount))
                    
                    Tax= {
                        "Type"           : "Tax",
                        "DocType"        : "",	
                        "Number"         : documentID,	
                        "Date"           : issueDate,	
                        "AccountNumber"  : AccountTax,	
                        "CompanyID"      : CompanyID,	
                        "Prefix"         : Prefix,	
                        "SupplierDoc"    : documentID.replace(Prefix,""),	
                        "DueDate"        : "",	
                        "Description"    : f"{invoiceLineTaxName} {float(checkEmpty(invoiceLinePercentage))}% S/ ${float(checkEmpty(invoiceLineTaxableAmount))}",	
                        "CostNumer"      : "",	
                        "Value"          : round(float(invoiceLineTaxAmount),2)       if invoiceLineTaxAmount      != ""    else invoiceLineTaxAmount,	
                        "Nature"         : "D",	
                        "TaxableAmount"  : round(float(invoiceLineTaxableAmount),2)   if invoiceLineTaxableAmount  != ""    else invoiceLineTaxableAmount,

                    }

                    if Tax["Value"] != 0 and Tax["Value"] != None and Tax["Value"] != "": 
                        AccountingDocs.append(Tax)
                    
                    
                    Document = {
                            "CustomerID"                 : CompanyID,
                            "CompanyID"                  : CCODE,	
                            "DocType"                    : "Invoice",	
                            "IssueDate"                  : issueDate,	
                            "ID"                         : documentID,
                            "Cufe"	                     : UUID,
                            "OrderNumber"                : OrderReference,	
                            "PaymentCode"                : PaymentMeansCode,	
                            "DueDate"                    : PaymentDueDate,	
                            "Prefix"                     : Prefix,	
                            "currencyID"                 : invoiceLinecurrencyID,	
                            "Rate"                       : round(float(CalculationRate),2)             if CalculationRate             != "" else CalculationRate,	
                            "InvoiceLine"                : invoiceLineID,	
                            "Description"                : invoiceLineItemDescription,	
                            "Subtotal"                   : round(float(invoiceLineExtensionAmount),2)  if invoiceLineExtensionAmount  != "" else invoiceLineExtensionAmount,	
                            "TaxID"                      : invoiceLineTaxID,	
                            "NombreImpuesto"             : invoiceLineTaxName,		
                            "Tax"                        : round(float(invoiceLineTaxAmount),2)        if invoiceLineTaxAmount        != "" else invoiceLineTaxAmount,	
                            "TaxableAmount"              : round(float(invoiceLineTaxableAmount),2)    if invoiceLineTaxableAmount    != "" else invoiceLineTaxableAmount,	
                            "TaxPercent"                 : round(float(invoiceLinePercentage),2)       if invoiceLinePercentage       != "" else invoiceLinePercentage,	
                            "IDWithHolding#1"            : "",	
                            "WithHoldingName#1"          : "",	
                            "WithHolding#1"              : "",	
                            "WithHoldingBaseAmount#1"    : "",	
                            "WithHoldingPercent#1"       : "",	
                            "IDWithHolding#2"            : "",	
                            "WithHoldingName#2"          : "",	
                            "WithHolding#2"              : "",	
                            "WithHoldingBaseAmount#2"    : "",
                            "WithHoldingPercent#2"       : ""

                            }
                    
                    invoiceWithholdingTaxTotals = invoiceLine.select("WithholdingTaxTotal")

                    if len(invoiceWithholdingTaxTotals)>0:
                            pass
                    else:
                        AccountNumbers = getAccountNumbersfromTable(RulesAccountNumbers,CompanyID,CCODE,"WithHolding",invoiceLineItemDescription,"ANY")
                        if AccountNumbers is not None:
                                WithHoldingAccount          = AccountNumbers["Account"]
                                CodeTax                     = AccountNumbers["CodeTax"]
                                Amount                      = round(float(invoiceLineExtensionAmount),2)       if invoiceLineExtensionAmount       != "" else invoiceLineExtensionAmount
                                
                                try:
                                    inoviceLineWithholdingName = CODIGO[CodeTax]
                                except KeyError:
                                    inoviceLineWithholdingName = CODIGO["0"]

                                
                                WithHolding = {
                                    "Type"           : "WithHolding",
                                    "DocType"        : "",	
                                    "Number"         : documentID,	
                                    "Date"           : issueDate,	
                                    "AccountNumber"  : WithHoldingAccount,	
                                    "CompanyID"      : CompanyID,	
                                    "Prefix"         : Prefix,	
                                    "SupplierDoc"    : documentID.replace(Prefix,""),	
                                    "DueDate"        : "",	
                                    "Description"    : f"{inoviceLineWithholdingName} {AccountNumbers['Percent']}% S/ ${Amount}",	
                                    "CostNumer"      : "",	
                                    "Value"          : Amount,	
                                    "Nature"         : "C",	
                                    "TaxableAmount"  : Amount,

                                }
                                AccountingDocs.append(WithHolding)

                        else:
                            WithHoldingAccount  = ""
                    
                    for count,invoiceWithholdingTaxTotal in enumerate(invoiceWithholdingTaxTotals,start=1):

                        if count > 2:
                            break

                        try:
                            inoviceLineWithholdingID                = invoiceWithholdingTaxTotal.select_one("TaxSubtotal>TaxCategory>TaxScheme>ID").text
                        except AttributeError:
                            inoviceLineWithholdingID                = ""
                        
                        
                        try:
                            inoviceLineWithholdingName              = invoiceWithholdingTaxTotal.select_one("TaxSubtotal>TaxCategory>TaxScheme>Name").text
                        except AttributeError:
                            inoviceLineWithholdingName              = ""
                        
                        
                        try:
                            inoviceLineWithholdingPercentage        = invoiceWithholdingTaxTotal.select_one("TaxSubtotal>TaxCategory>Percent").text
                        except AttributeError:
                            inoviceLineWithholdingPercentage        = ""
                        
                        
                        try:
                            invoiceLineWithholdingTaxAmount         = invoiceWithholdingTaxTotal.select_one("TaxSubtotal>TaxAmount").text
                        except AttributeError:
                            invoiceLineWithholdingTaxAmount         = ""
                        
                        
                        try:
                            invoiceLineWithholdingTaxableAmount     = invoiceWithholdingTaxTotal.select_one("TaxSubtotal>TaxableAmount").text
                        except AttributeError:
                            invoiceLineWithholdingTaxableAmount     = ""
                        
                        Document[f"IDWithHolding#{count}"]                     = inoviceLineWithholdingID
                        Document[f"WithHoldingName#{count}"]	               = inoviceLineWithholdingName
                        Document[f"WithHolding#{count}"]	                   = round(float(invoiceLineWithholdingTaxAmount),2)     if invoiceLineWithholdingTaxAmount     != "" else invoiceLineWithholdingTaxAmount 
                        Document[f"WithHoldingBaseAmount#{count}"]	           = round(float(invoiceLineWithholdingTaxableAmount),2) if invoiceLineWithholdingTaxableAmount != "" else invoiceLineWithholdingTaxableAmount
                        Document[f"WithHoldingPercent#{count}"]                = round(float(inoviceLineWithholdingPercentage),2)    if inoviceLineWithholdingPercentage    != "" else inoviceLineWithholdingPercentage

                        AccountNumbers = getAccountNumbersfromTable(RulesAccountNumbers,CompanyID,CCODE,"WithHolding",invoiceLineItemDescription,float(checkEmpty(inoviceLineWithholdingPercentage)))
                        if AccountNumbers is not None:
                            WithHoldingAccount  = AccountNumbers["Account"]
                        else:
                            WithHoldingAccount  = ""
                        
                        WithHoldingCalculation = (float(checkEmpty(inoviceLineWithholdingPercentage)) / 100) * float(checkEmpty(invoiceLineWithholdingTaxableAmount))
                    
                        WithHolding = {
                            "Type"           : "WithHolding",
                            "DocType"        : "",	
                            "Number"         : documentID,	
                            "Date"           : issueDate,	
                            "AccountNumber"  : WithHoldingAccount,	
                            "CompanyID"      : CompanyID,	
                            "Prefix"         : Prefix,	
                            "SupplierDoc"    : documentID.replace(Prefix,""),	
                            "DueDate"        : "",	
                            "Description"    : f"{inoviceLineWithholdingName} {float(checkEmpty(inoviceLineWithholdingPercentage))}% S/ ${float(checkEmpty(invoiceLineWithholdingTaxableAmount))}"                if WithHoldingCalculation               != ""    else WithHoldingCalculation,	
                            "CostNumer"      : "",	
                            "Value"          : round(float(invoiceLineWithholdingTaxAmount),2)       if invoiceLineWithholdingTaxAmount      != ""    else invoiceLineWithholdingTaxAmount,	
                            "Nature"         : "C",	
                            "TaxableAmount"  : round(float(invoiceLineWithholdingTaxableAmount),2)   if invoiceLineWithholdingTaxableAmount  != ""    else invoiceLineWithholdingTaxableAmount,

                        }
                        AccountingDocs.append(WithHolding)
                        WithHoldingTotalTax+=float(checkEmpty(invoiceLineWithholdingTaxAmount))
                    
                    InvoiceLineTotalTax+=float(checkEmpty(invoiceLineTaxAmount))
                    InvoiceLineExtensionTotalAmount+=float(checkEmpty(invoiceLineExtensionAmount))
                    Documents.append(Document)
        
                TotalPayable = float(checkEmpty(InvoiceLineExtensionTotalAmount)) + float(checkEmpty(InvoiceLineTotalTax)) - float(checkEmpty(WithHoldingTotalTax))
                Payable = {
                        "Type"           : "Payables",
                        "DocType"        : "",	
                        "Number"         : documentID,	
                        "Date"           : issueDate,	
                        "AccountNumber"  : "",	
                        "CompanyID"      : CompanyID,	
                        "Prefix"         : Prefix,	
                        "SupplierDoc"    : documentID.replace(Prefix,""),	
                        "DueDate"        : PaymentDueDate,	
                        "Description"    : f"Valor a pagar según {documentID}",	
                        "CostNumer"      : "",	
                        "Value"          : round(float(TotalPayable),2) if TotalPayable  != ""  else TotalPayable,	
                        "Nature"         : "C",	
                        "TaxableAmount"  : "",

                                }
                AccountingDocs.append(Payable)
            
    return True

def getService():

    SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
    CREDENTIALS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),'creds.json')
    
    creds = None

    if os.path.exists(os.path.join(os.path.dirname(os.path.abspath(__file__)),'token.json')):
        creds = Credentials.from_authorized_user_file(os.path.join(os.path.dirname(os.path.abspath(__file__)),'token.json'), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)),'token.json'), 'w') as token:
            token.write(creds.to_json())

    service = build('gmail', 'v1', credentials=creds)
    return service


def searchMessages(service, user_id, sender, subject):
    query = f'subject:{subject}'
    response = service.users().messages().list(userId=user_id, q=query).execute()
    messages = response.get('messages', [])
    
    return messages


def getLatestEmail(service, sender, subject):
    
    messages = searchMessages(service, 'me', sender, subject)
    if not messages:
        return (None,None)
    
    latestMessageID = messages[0]['id']
    message         = service.users().messages().get(userId='me', id=latestMessageID).execute()
    headers         = message['payload']['headers']
    dateHeader      = next(header['value'] for header in headers if header['name'] == 'Date')
    timeTuple       = parsedate_tz(dateHeader)
    timesTamp       = mktime_tz(timeTuple)
    dateOBJ         = datetime.fromtimestamp(timesTamp)
    dateOBJ         = dateOBJ.astimezone(timezone.utc)
    dateOBJ         = dateOBJ.replace(tzinfo=None)
    
    messagePayload  = message.get('payload')

    def getMIMEparts(parts, mimeType):
        for part in parts:
            if part.get('mimeType') == mimeType:
                return part
            if part.get('parts'):
                return getMIMEparts(part['parts'], mimeType)
        return None

    if 'parts' in messagePayload:
        html_part = getMIMEparts(messagePayload['parts'], 'text/html')
        if html_part and 'data' in html_part['body']:
            bodyData = base64.urlsafe_b64decode(html_part['body']['data']).decode('utf-8')
        else:
            bodyData = "No HTML body found."
    else:
        bodyData =  base64.urlsafe_b64decode(messagePayload['body']['data']).decode('utf-8')

    return (bodyData,dateOBJ)


def getURLFromGmail(companyCode,computerTime):
    retryCount = 0
    while True:
        service = getService()
        sender = 'face4-vp@dian.gov.co'
        subject = 'Estimado (a),'
        emailBody,dateOBJ = getLatestEmail(service, sender, subject)
        
        if emailBody and dateOBJ > computerTime:

            soup =  BeautifulSoup(emailBody,"html.parser")
            
            try:
                extractedLink = soup.find("a",{"title":"ingreso"})["href"]
            except (KeyError,TypeError):
                extractedLink = None
            
            
            if extractedLink:
                pattern = rf"rk={companyCode}(&|$)"
                isValid = bool(re.search(pattern, extractedLink))
                if isValid:
                    return extractedLink
                else:
                    retryCount+=1
                    if retryCount>20:
                        return ""
                    time.sleep(5)
            else:
                    retryCount+=1
                    if retryCount>20:
                        return ""
                    time.sleep(5)
        else:
            retryCount+=1
            if retryCount>20:
                return ""
            time.sleep(5)


def sendEmail(UserCode,CompanyCode,IDType,__RequestVerificationToken,RecaptchaToken,cookies):
    while True:
        headers = {
        'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:121.0) Gecko/20100101 Firefox/121.0',
        'Accept': '*/*',
        'Accept-Language': 'en-US,en;q=0.5',
        'Referer': 'https://catalogo-vpfe.dian.gov.co/',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'X-Requested-With': 'XMLHttpRequest',
        'Origin': 'https://catalogo-vpfe.dian.gov.co',
        'DNT': '1',
        'Sec-GPC': '1',
        'Connection': 'keep-alive',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        }

        data = {
            '__RequestVerificationToken': __RequestVerificationToken,
            'RecaptchaToken': RecaptchaToken ,
            'IdentificationType': IDType,
            'UserCode': UserCode,
            'CompanyCode': CompanyCode,
            'X-Requested-With': 'XMLHttpRequest',
        }

        response = requests.post(
            'https://catalogo-vpfe.dian.gov.co/User/CompanyAuthentication',
            cookies=cookies,
            headers=headers,
            data=data,
        )

        if response.status_code == 200:

            jsonData = response.json()
            # print(jsonData)
            if jsonData["Code"] == 200:
                return {"status":"Success","status_code":response.status_code,"detail":response.json()}
            else:
                continue
                # return {"status":"Error","status_code":response.status_code,"detail":response.json()}
    
        else:
            continue
            # return {"status":"Error","status_code":response.status_code,"detail":response.json()}


def parseCookiesV1(cookies):
    
    requestCookies = {
    '__RequestVerificationToken': '',
    'ARRAffinity': '',
    'ARRAffinitySameSite': ''
    }

    for cookie in cookies:
        if cookie["name"]=="ARRAffinity":
            requestCookies["ARRAffinity"]=cookie["value"]
        
        elif cookie["name"]=="ARRAffinitySameSite":
            requestCookies["ARRAffinitySameSite"]=cookie["value"]
        
        elif cookie["name"]=="__RequestVerificationToken":
            requestCookies["__RequestVerificationToken"]=cookie["value"]

    
    return requestCookies
 


def sendEmailtoGmail(playwright,UserCode,CompanyCode,IDType):

    try:
    
        websiteURL = "https://catalogo-vpfe.dian.gov.co/User/CompanyLogin"

        browser    = playwright.firefox.launch(headless=True)
        context    = browser.new_context()
        
        page       = context.new_page()

        page.goto(websiteURL)     
        

        __RecaptchaToken = ""
        
        while __RecaptchaToken == "":
            
            pageContent = page.content()
            soup        = BeautifulSoup(pageContent,"html.parser")
            
            __RequestVerificationToken = soup.find("form",{"id":"form0"}).find("input",{"name":"__RequestVerificationToken"})["value"]
            __RecaptchaToken           = soup.find("form",{"id":"form0"}).find("input",{"name":"RecaptchaToken"})["value"]

        
        cookies = context.cookies()
        
        getCookies = parseCookiesV1(cookies)
        

        response = sendEmail(UserCode,CompanyCode,IDType,__RequestVerificationToken,__RecaptchaToken,getCookies)
        if response["status"] == "Success":
            pass 
        else:
            raise Exception(response["detail"])

        browser.close()

    except Exception as E:
        return {"status":"Error","detail":E}
    else:
        return {"status":"Success","detail":"Done :)"}


def getDateInput():
    dateRegex = r'^(\d{4})\/(0[1-9]|1[0-2])\/(0[1-9]|[12][0-9]|3[01])$'
    while True:
        dateInput = input("\t\tEnter the date in YYYY/MM/DD format: ")
        if re.match(dateRegex, dateInput):
            try:
                datetime.strptime(dateInput, '%Y/%m/%d')
                return dateInput
            
            except ValueError:
                print("Invalid date, please enter a valid date in YYYY/MM/DD format.")
        else:
            print("Invalid format, please enter a date in YYYY/MM/DD format.")



def downloadFile(trackId,cookies,CompanyCode):
    
    
    DOWNLOADED_FILE_PATH              = os.path.join(os.path.dirname(os.path.abspath(__file__)),"Downloaded Files")
    DOWNLOADED_FILE_COMPANY_CODE_PATH = os.path.join(DOWNLOADED_FILE_PATH,CompanyCode)

    if not os.path.exists(DOWNLOADED_FILE_COMPANY_CODE_PATH):
        os.makedirs(DOWNLOADED_FILE_COMPANY_CODE_PATH)

    headers = {
        'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:121.0) Gecko/20100101 Firefox/121.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
        'Referer': 'https://catalogo-vpfe.dian.gov.co/',
        'DNT': '1',
        'Sec-GPC': '1',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
    }

    params = {
        'trackId': trackId,
    }

    response = requests.get(
        'https://catalogo-vpfe.dian.gov.co/Document/DownloadZipFiles',
        params=params,
        cookies=cookies,
        headers=headers,
    )

    if response.status_code == 200:

        fileName = os.path.join(DOWNLOADED_FILE_COMPANY_CODE_PATH,f"{trackId}.zip") 
        with open(fileName, 'wb') as file:
            file.write(response.content)
        
        return {"status":"Success","detail":f"{fileName} downloaded successfully :)"}
    else:
        return {"status":"Error","detail":f"Failed to download file: status code {response.status_code}"}


def parseCookiesV2(cookies):
    
    requestCookies = {
        '.AspNet.ApplicationCookie': '',
        'ARRAffinity': '',
        'ARRAffinitySameSite': '',
        'ASP.NET_SessionId': '',
        '__RequestVerificationToken': '',
    }

    for cookie in cookies:
        if cookie["name"]=="ARRAffinity":
            requestCookies["ARRAffinity"]=cookie["value"]
        
        elif cookie["name"]=="ARRAffinitySameSite":
            requestCookies["ARRAffinitySameSite"]=cookie["value"]
        
        elif cookie["name"]=="__RequestVerificationToken":
            requestCookies["__RequestVerificationToken"]=cookie["value"]
        
        elif cookie["name"]==".AspNet.ApplicationCookie":
            requestCookies[".AspNet.ApplicationCookie"]=cookie["value"]
        
        elif cookie["name"]=="ASP.NET_SessionId":
            requestCookies["ASP.NET_SessionId"]=cookie["value"]

    
    return requestCookies


def DownloadZipFiles(playwright,websiteURL,CompanyCode,startDateRange,endDateRange,documentsURL):
    try:

        browser    = playwright.firefox.launch(headless=True)
        context    = browser.new_context()
        
        page       = context.new_page()

        
        page.goto(websiteURL)


        print(f"going to URL={documentsURL}")

        page.goto(documentsURL)
        

        DateRangePickerInput = "//input[@id='dashboard-report-range']"
        

        page.wait_for_selector(f'xpath={DateRangePickerInput}', state="visible")
        page.click(f'xpath={DateRangePickerInput}')    
        
        

        DateRangePickerStartInput = "//input[@name='daterangepicker_start']"

        page.wait_for_selector(f'xpath={DateRangePickerStartInput}', state="visible")
        page.fill(f'xpath={DateRangePickerStartInput}', "")
        page.type(DateRangePickerStartInput, startDateRange, delay=100)
        page.press(DateRangePickerStartInput, "Enter")

        DateRangePickerEndInput   = "//input[@name='daterangepicker_end']"

        page.wait_for_selector(f'xpath={DateRangePickerEndInput}', state="visible")
        page.fill(f'xpath={DateRangePickerEndInput}', "")
        page.type(DateRangePickerEndInput, endDateRange, delay=100)
        page.press(DateRangePickerEndInput, "Enter")
        
        nextInput  = "//*[@id='DocumentKey']"

        page.wait_for_selector(f'xpath={nextInput}', state="visible")
        page.click(f'xpath={nextInput}')

        submitButton = "//button[text()='Buscar']"

        page.wait_for_selector(f'xpath={submitButton}', state="visible")
        page.click(f'xpath={submitButton}')


        cookies = context.cookies()
        
        getCookies = parseCookiesV2(cookies)

        while True:    

            tableDocumentsTable = "//*[@id='tableDocuments']/tbody/tr"
            page.wait_for_selector(f'xpath={tableDocumentsTable}', state="visible")

            getAllRows = page.query_selector_all(f'xpath={tableDocumentsTable}')

            for row in getAllRows:

                trackIDButton  = "./td/button"
                trackIDElement = row.query_selector(f'xpath={trackIDButton}')
                try:
                    trackID        = trackIDElement.get_attribute("data-id")
                    response = downloadFile(trackID,getCookies,CompanyCode)
                    if response["status"] == "Success":
                        print(response)
                        continue
                    else:
                        raise Exception(response["detail"]) 
                except AttributeError:
                    pass



            nextPage                 = "//*[@id='tableDocuments_next']"
            nextPageElement          = page.query_selector(f'xpath={nextPage}')
            nextPageAttributeClass   = nextPageElement.get_attribute("class")

            if "disabled" in nextPageAttributeClass:
                break
            else:
                page.click(f'xpath={nextPage}')


        browser.close()
    except Exception as E:
        return {"status":"Error","detail":E}
    else:
        return {"status":"Success","detail":"Done :)"}

def unZipAndOrganize(folderPath):
    for mainItem in os.listdir(folderPath):
        mainItemPath = os.path.join(folderPath, mainItem)

        if os.path.isdir(mainItemPath):
            for item in os.listdir(mainItemPath):
                if item.endswith('.zip'):
                    zipPath = os.path.join(mainItemPath, item)
                    try:
                        with zipfile.ZipFile(zipPath, 'r') as zipRef:
                            zipRef.extractall(mainItemPath)
                    except Exception as E:
                        pass
                    os.remove(zipPath)

            for item in os.listdir(mainItemPath):
                if not os.path.isdir(os.path.join(mainItemPath, item)):
                    fileType = item.split('.')[-1]
                    destinationFolder = os.path.join(mainItemPath, fileType)

                    if not os.path.exists(destinationFolder):
                        os.makedirs(destinationFolder)

                    destinationPath = os.path.join(destinationFolder, item)

                    newDestinationPath = destinationPath
                    baseName, extension = os.path.splitext(item)
                    counter = 1
                    while os.path.exists(newDestinationPath):
                        newFileName = f"{baseName}_{counter}{extension}"
                        newDestinationPath = os.path.join(destinationFolder, newFileName)
                        counter += 1

                    shutil.move(os.path.join(mainItemPath, item), newDestinationPath)

def main(computerTime):
    with open(os.path.join(os.path.dirname(os.path.abspath(__file__)),"UserandCompanyCodes.txt"),"r") as R:
        
        print("Start Date Range:")
        startDateRange         = getDateInput()
        
        print("End Date Range:")
        endDateRange           = getDateInput()
        
        for count,i in enumerate(R.readlines(),start=1):

            if i == "\n" or i == " ":
                continue

            UserCode    = i.split("|")[0].strip()
            CompanyCode = i.split("|")[1].strip()
            IDType      = i.split("|")[2].replace("\n","").strip()

            print("=======================================")

            print(f"{count})")
            print(f"\tUserCode    : {UserCode}   ")
            print(f"\tCompanyCode : {CompanyCode}\n")

            print("=======================================")

        
            with sync_playwright() as playwright:
                
                print("Sending email...")
                response = sendEmailtoGmail(playwright,UserCode,CompanyCode,getIDType(IDType))
                if response["status"] == "Success":
                    print("Email sent successfully.")
                    print("Waiting for the Email ....")
                    websiteURL = getURLFromGmail(CompanyCode,computerTime)
                    if websiteURL == "":
                        print("Email not received , therefore skipping this CompanyCode...")
                        continue
                    print("Email received successfully.")
                    print(websiteURL)
                    print("Downloading the zip files...")

                    receivedDocumentsURL   = "https://catalogo-vpfe.dian.gov.co/Document/Received"

                    response = DownloadZipFiles(playwright,websiteURL,CompanyCode,startDateRange,endDateRange,receivedDocumentsURL)
                    if response["status"] == "Success":
                        
                        sentDocumentsURL   = "https://catalogo-vpfe.dian.gov.co/Document/Sent"

                        response = DownloadZipFiles(playwright,websiteURL,CompanyCode,startDateRange,endDateRange,sentDocumentsURL)
                        if response["status"] == "Success":
                            pass
                        else:
                            return {"status":"Error","detail":response["detail"]}
                    else:
                        return {"status":"Error","detail":response["detail"]}
                else:
                            return {"status":"Error","detail":response["detail"]}
        else:
            return {"status":"Success","detail":"Done :)"}

if __name__ == '__main__':
    
    utcTime = datetime.now(timezone.utc)
    utcTime = utcTime.replace(tzinfo=None)
    
    response = main(utcTime)
    print(response)
    
    baseDirectoryPath = os.path.dirname(os.path.abspath(__file__))
    
    unZipAndOrganize(os.path.join(baseDirectoryPath,"Downloaded Files"))

    for countryCode in os.listdir(os.path.join(baseDirectoryPath,"Downloaded Files")):

        print(countryCode)

        invoiceLines = parsingXMLFiles(os.path.join(os.path.join(os.path.join(baseDirectoryPath,"Downloaded Files"),countryCode),"xml"),countryCode)
        writeDictstoExcel(invoiceLines,os.path.join(os.path.join(os.path.join(baseDirectoryPath,"Downloaded Files"),countryCode),"InvoiceLineTemplate.xlsx"))
        Parties             = []
        Documents           = []
        AccountingDocs      = []
    logsFilePath = os.path.join(os.path.join(baseDirectoryPath,"Logs"),f"main.log")
    with open(logsFilePath,"w") as L:
        L.write(f"{str(datetime.now())}\n\n\n\n{response}")

