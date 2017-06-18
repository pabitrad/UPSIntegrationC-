//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop

#include "UPSInterface.h"
#include <WinInetIntf.h>
#include <msxmldom.hpp>
#include <XMLDoc.hpp>
#include <xmldom.hpp>
#include <XMLIntf.hpp>
#include "Functions.h"
#include "resources.h"
#include "Logger.h"
#include "CountryCodes.h"
#include <InvokeRegistry.hpp>
#include <Soaphttpclient.hpp>
//---------------------------------------------------------------------------
#pragma package(smart_init)
using namespace asap;
//---------------------------------------------------------------------------
TUPSInterface UPSInterface;
//---------------------------------------------------------------------------
void TUPSInterface::InitMyOBServiceMap()
{
  MyOBServiceMap[ "UPS 2nd Day Air" ] = TUPSServiceCode( "02", "2DA" );
  MyOBServiceMap[ "UPS 2nd Day Air AM" ] = TUPSServiceCode( "59", "2DM" );
  MyOBServiceMap[ "UPS EXPRESS" ] = TUPSServiceCode( "07","01", true );
  MyOBServiceMap[ "UPS GROUND" ] = TUPSServiceCode( "03", "GND" );
  MyOBServiceMap[ "UPS 3 Day Select" ] = TUPSServiceCode( "12", "3DS" );
  MyOBServiceMap[ "UPS Next Day Air" ] = TUPSServiceCode( "01", "1DA" );
  MyOBServiceMap[ "UPS NDA Early AM" ] = TUPSServiceCode( "14", "1DM", true );
  MyOBServiceMap[ "UPS NDA Saturday" ] = TUPSServiceCode( "01", "1DMS", false, true );
  MyOBServiceMap[ "UPS RED SAVER" ] = TUPSServiceCode( "13", "1DP" );
  MyOBServiceMap[ "UPS STANDARD" ] = TUPSServiceCode( "11", "03" );
  MyOBServiceMap[ "UPS World Exp +" ] = TUPSServiceCode( "54", "21", true );
  MyOBServiceMap[ "UPS World Expedited" ] = TUPSServiceCode( "08", "05", true );

  MyOBServiceMap[ "UPS COLLECT BLUE" ] = MyOBServiceMap[ "UPS BLUE" ];
  MyOBServiceMap[ "UPS COLLECT GRD" ] = MyOBServiceMap[ "UPS GROUND" ];
  MyOBServiceMap[ "UPS COLLECT ORANGE" ] = MyOBServiceMap[ "UPS ORANGE" ];
  MyOBServiceMap[ "UPS COLLECT RED" ] = MyOBServiceMap[ "UPS RED" ];
  MyOBServiceMap[ "UPS COLLECT REDSAVER" ] = MyOBServiceMap[ "UPS RED SAVER" ];
}
//---------------------------------------------------------------------------
bool TUPSInterface::MyOBServiceToUPSServiceCode( const String& PMyOBService, TUPSServiceCode &PServiceCode )
{
  std::map<String, TUPSServiceCode>::const_iterator it;

  it = MyOBServiceMap.find( PMyOBService );
  if( it == MyOBServiceMap.end() )
    return false;

  PServiceCode = it->second;
  return true;
}
//---------------------------------------------------------------------------
bool TUPSInterface::RetrieveCredentials( String &URL, String &Licence,
  String &UserName, String &Password )
{
  Licence = IniFile->ReadString( "UPS", "AccessLicence", "" );
  UserName = IniFile->ReadString( "UPS", "UserId", "" );
  Password = IniFile->ReadString( "UPS", "Password", "" );

  if( Licence.Length() == 0 || UserName.Length() == 0 || Password.Length() == 0 )
    return false;

  if( IniFile->ReadBool( "UPS", "Testing", true ) )
    URL = "https://wwwcie.ups.com/";
  else
    URL = "https://onlinetools.ups.com/";

  return true;
}
//---------------------------------------------------------------------------
bool TUPSInterface::UPSGetRate( TUPSServiceCode &PServiceType, double const PWeightPounds,
  double const PWeightOunces, double const PkgLength, double const PkgWidth, double const PkgHeight,
  TAddress const &PFromAddress, TAddress const &PToAddress, String &PEstimatedCost, String &PTimeCommit, 
  String &AddressType, String &PErrorMessage )
{
//	ParentWindow = PParentWindow;
  PEstimatedCost = "0.00";
  PErrorMessage = "Unknown Error";
  String AccessLicenceNo;
  String AccessUserId;
  String AccessPassword;
  String Url;
  String BaseUrl;

  if( !RetrieveCredentials( BaseUrl, AccessLicenceNo, AccessUserId, AccessPassword ) )
  {
    MsgBox( "UPS rate access not setup. Please configure before querying rate information.", MB_ICONERROR | MB_OK );
    return false;
  }
  CheckError( !BaseUrl.IsEmpty(), "UPS Url is empty in configuration file" );

  // Append the particular service we're using to the URL
  Url = BaseUrl + "ups.app/xml/Rate";

  TReplaceFlags rf;
  rf << rfReplaceAll;

  String AccessXML = LoadResourceDataString( UPS_ACCESS );
  String RateRequest = LoadResourceDataString( UPS_RATE );

  //String AddressType;

  AccessXML = StringReplace( AccessXML, "{ACCESSLICENCENUMBER}", AccessLicenceNo, rf );
  AccessXML = StringReplace( AccessXML, "{USERID}", AccessUserId, rf );
  AccessXML = StringReplace( AccessXML, "{PASSWORD}", AccessPassword, rf );

  GetAddressType(BaseUrl, AccessXML, PToAddress, PErrorMessage, AddressType);

  double Weight;

  Weight = ( ( PWeightPounds * 16.0f ) + PWeightOunces ) / 16.0f;

  RateRequest = StringReplace( RateRequest, "{shiptocity}", PToAddress.City, rf );
  RateRequest = StringReplace( RateRequest, "{shiptozip}", PToAddress.Zip, rf );
  RateRequest = StringReplace( RateRequest, "{shiptocountry}", GetCountryCodeFromName( PToAddress.Country ), rf );
  RateRequest = StringReplace( RateRequest, "{shipfromzip}", PFromAddress.Zip, rf );
  RateRequest = StringReplace( RateRequest, "{shipfromcountry}", GetCountryCodeFromName( PFromAddress.Country ), rf );
  RateRequest = StringReplace( RateRequest, "{length}", FloatToStrF( PkgLength, ffFixed, 15, 2 ), rf );
  RateRequest = StringReplace( RateRequest, "{width}", FloatToStrF( PkgWidth, ffFixed, 15, 2 ), rf );
  RateRequest = StringReplace( RateRequest, "{height}", FloatToStrF( PkgHeight, ffFixed, 15, 2 ), rf );
  RateRequest = StringReplace( RateRequest, "{weight}", FloatToStrF( Weight, ffFixed, 15, 1 ), rf );
  RateRequest = StringReplace( RateRequest, "{service}", PServiceType.RateCode, rf );

  if( PServiceType.ShipSaturday )
    RateRequest = StringReplace( RateRequest, "{saturdaypickup}", "<SaturdayPickupIndicator />", rf );
  else
    RateRequest = StringReplace( RateRequest, "{saturdaypickup}", "", rf );

  String PDataOut;
  String Input;
  String Cost;
  String CurrencyCode;
  TWinInet InetIntf;
  _di_IXMLDocument xml;
  _di_IXMLNode TempNode;

  /*THTTPRIO *HttpRio = new THTTPRIO( NULL );
  HttpRio->OnBeforeExecute = SoapBeforeExecute;
  HttpRio->OnAfterExecute = SoapAfterExecute;
  HttpRio->QueryInterface( InetIntf );
  */
  LOG( "sending rate request to UPS" );

  try {
    //HttpRio->URL = Url;
    PDataOut = InetIntf.Post( Url, AccessXML + RateRequest );
  }

  catch( EPropertyConvertError const &E ) {
    LOGEXCEPTION( E );
    PErrorMessage = E.Message;
  //	throw EASAPException( E.Message );
  }

  catch( ERemotableException const &E ) {
    LOGEXCEPTION( E );
    PErrorMessage = E.Message;
  //	throw EASAPException( E.Message );
  }

  catch( Exception const &E ) {
    LOGEXCEPTION(E);
    MsgBox( "Failed to send rate request to UPS.", MB_ICONERROR | MB_OK );
    PErrorMessage = E.Message;
    return false;
    }

  LOG( "received rate response" );

  try {
    xml = NewXMLDocument();
    xml->LoadFromXML( PDataOut );
        xml->SaveToFile("ups_soap_responce.xml");
    _di_IXMLNode ResponseNode = xml->DocumentElement->ChildNodes->FindNode( "Response" );
    if( !ResponseNode )
      throw Exception( "can't find Response node." );

    TempNode = ResponseNode->ChildNodes->FindNode( "ResponseStatusCode" );
    if( !TempNode )
      throw Exception( "can't find ResponseStatusCode node." );

    int StatusCode = StrToInt( TempNode->Text );

    TempNode = ResponseNode->ChildNodes->FindNode( "ResponseStatusDescription" );
    if( !TempNode )
      throw Exception( "can't find ResponseStatusDescription node." );

    String StatusDesc = TempNode->Text;

    if( StatusCode != 1 )
    {
      LOGX( "UPS returned error (%d): %s", ( StatusCode, StatusDesc ) );
      TempNode = ResponseNode->ChildNodes->FindNode( "Error" );
      if( !TempNode )
      throw Exception( "can't find Error node." );
      TempNode = TempNode->ChildNodes->FindNode( "ErrorDescription" );
      if( !TempNode )
      throw Exception( "can't find ErrorDescription node." );
      PErrorMessage = TempNode->Text;
      xml->SaveToFile("ups_soap_bad_response.xml");
      //MsgBox( "UPS returned an error code from the request. Rate query failed.", MB_ICONERROR | MB_OK );
      return false;
    }

    _di_IXMLNode RatedShipmentNode = xml->DocumentElement->ChildNodes->FindNode( "RatedShipment" );
    if( !ResponseNode )
      throw Exception( "can't find RatedShipment node." );

    _di_IXMLNode TotalChargesNode = RatedShipmentNode->ChildNodes->FindNode( "TotalCharges" );
    if( !TotalChargesNode )
      throw Exception( "can't find TotalChargesNode node." );

    /*_di_IXMLNode GuaranteedDaysToDeliveryNode = RatedShipmentNode->ChildNodes->FindNode( "GuaranteedDaysToDelivery" );
    if( !GuaranteedDaysToDeliveryNode )
      throw Exception( "can't find GuaranteedDaysToDeliveryNode node." );
    if( !GuaranteedDaysToDeliveryNode->Text.IsEmpty( ) )
       PTimeCommit = GuaranteedDaysToDeliveryNode->Text + " day(s).";*/
    
    if (UPSGetDeliveryTime (PServiceType, BaseUrl, AccessXML, PFromAddress, PToAddress, PErrorMessage, PTimeCommit)== false)
    {
        LOG("Can't find delivery date");
    }

    _di_IXMLNode ScheduledDeliveryTimeNode = RatedShipmentNode->ChildNodes->FindNode( "ScheduledDeliveryTime" );
    if( !ScheduledDeliveryTimeNode )
      throw Exception( "can't find ScheduledDeliveryTime node." );
    if( !ScheduledDeliveryTimeNode->Text.IsEmpty() )
      PTimeCommit = PTimeCommit + " Sched. at " + ScheduledDeliveryTimeNode->Text;

    _di_IXMLNode CurrencyCodeNode = TotalChargesNode->ChildNodes->FindNode( "CurrencyCode" );
    if( !CurrencyCodeNode )
      throw Exception( "can't find CurrencyCodeNode node." );

    CurrencyCode = CurrencyCodeNode->Text;

    TempNode = TotalChargesNode->ChildNodes->FindNode( "MonetaryValue" );
    if( !TempNode )
      throw Exception( "can't find MonetaryValue node." );

    Cost = TempNode->Text;
  }

  catch( Exception const &E ) {
    LOGX( "UPS xml invalid, %s", ( E.Message ) );
    MsgBox( "The XML response from UPS was invalid.", MB_ICONERROR | MB_OK );
    return false;
  }

  if( CurrencyCode != "USD" )
  {
    LOGX( "UPS returned unsupported currency code '%s'", ( CurrencyCode ) );
    MsgBox( "UPS returned an unsupported currency code '" + CurrencyCode + "'", MB_ICONERROR | MB_OK );
    return false;
  }

  PEstimatedCost = Cost;
  PErrorMessage = "";
  return true;
}
//---------------------------------------------------------------------------
void __fastcall TUPSInterface::SoapBeforeExecute(const String MethodName, WideString &SOAPRequest)
{

}
//---------------------------------------------------------------------------
void __fastcall TUPSInterface::SoapAfterExecute(const String MethodName, Classes::TStream* SOAPResponse)
{

}
//---------------------------------------------------------------------------
bool TUPSInterface::UPSGetRate( TUPSServiceCode &PServiceType, double const PWeightPounds,
  double const PWeightOunces, double const PkgLength, double const PkgWidth, double const PkgHeight,
  TAddress const &PFromAddress, TAddress const &PToAddress, TXmlNode *PRowNode )
{
  String NotificationMessage;
  String EstRate = "";
  String TimeCommit;
  String AddressType;
  
  UPSGetRate( PServiceType, PWeightPounds, PWeightOunces, PkgLength, PkgWidth, PkgHeight,
    PFromAddress, PToAddress, EstRate, TimeCommit, AddressType, NotificationMessage );

  std::map<String, TUPSServiceCode>::iterator it = MyOBServiceMap.begin();
  String ServiceName;
  while( it != MyOBServiceMap.end() )
  {
     if( it->second.RateCode == PServiceType.RateCode )
     {
       ServiceName = it->first;
       break;
     }
     it++;
  }

  PRowNode->WriteAttributeString( "PostName", "UPS" );
  PRowNode->WriteAttributeString( "ServiceName", ServiceName );
  PRowNode->WriteAttributeString( "Cost", EstRate );
  PRowNode->WriteAttributeString( "TimeCommit", TimeCommit );
  PRowNode->WriteAttributeString( "Notification", NotificationMessage );
  PRowNode->WriteAttributeString( "MyCost", EstRate );
  PRowNode->WriteAttributeString( "ServiceCode", PServiceType.RateCode );
  PRowNode->WriteAttributeString( "AddressType", AddressType);
}
//---------------------------------------------------------------------------
void TUPSInterface::CalculateUPS( double PoundWeight, double OunceWeight, double PkgLength,
   double PkgWidth, double PkgHeight, TAddress FromAddr, TAddress ToAddr, TXmlNode *PListNode, String PServiceCode )
{
  try {
    String EstRate;
    CheckError( !( PoundWeight == 0.00f && OunceWeight == 0.00f ),
       "Weight of package must be valid numeric value greater than 0." );
    CheckError( !FromAddr.Zip.IsEmpty(), "From Zip is Empty" );
   //	CheckError( !ToAddr.City.IsEmpty(), "To City is Empty" );
    CheckError( !ToAddr.Zip.IsEmpty(), "To Zip is Empty" );
    CheckError( !ToAddr.Country.IsEmpty(), "To Country is Empty" );

    String Progress = "Please wait. Contacting UPS Rate services.";
    String CountryCode = GetCountryCodeFromName( ToAddr.Country );
    if( PServiceCode != "-1" )
    {
      TXmlNode *Row = NULL;
      Row = PListNode->NodeNew( "row" );
      TUPSServiceCode UPSService;
      UPSService.RateCode = PServiceCode;
      UPSService.TimeInTransitCode = GetTimeInTransitCode(PServiceCode);
      CheckError(!UPSService.TimeInTransitCode.IsEmpty(), "TimeInTransitCode is empty");

      UPSGetRate( UPSService, PoundWeight, OunceWeight, PkgLength, PkgWidth, PkgHeight,
        FromAddr, ToAddr, Row );
    }
    else
    {
      std::map<String, TUPSServiceCode>::iterator it = MyOBServiceMap.begin();
      while( it != MyOBServiceMap.end() )
      {
        bool DoGetRate = true;
        if( CountryCode == "US" )
        {
          DoGetRate = ( !it->second.International );
        }
        else
        {
          DoGetRate = ( it->second.International );
        }

        if( DoGetRate )
        {
          DoGetRate = !( ( it->first == "UPS BLUE" ) ||
          ( it->first== "UPS COLLECT BLUE" ) || ( it->first == "UPS COLLECT ORANGE" ) ||
          ( it->first == "UPS COLLECT RED" )  || ( it->first == "UPS ORANGE" ) ||
          ( it->first == "UPS RED" ) );
        }
        if( DoGetRate )
        {
          TXmlNode *Row = NULL;
          Row = PListNode->NodeNew( "row" );
          UPSGetRate( it->second, PoundWeight, OunceWeight, PkgLength, PkgWidth, PkgHeight,
            FromAddr, ToAddr, Row );
        }
        it++;
      }
    }
  }

  catch( Exception const &E ) {
    LOGEXCEPTION( E );
  }
}

bool TUPSInterface::UPSGetDeliveryTime (TUPSServiceCode const &PServiceType, String const& BaseUrl, String const& AccessXML, TAddress const &PFromAddress, TAddress const &PToAddress, String &PErrorMessage, 
                                       String &PTimeCommit)
{
    String TimeInTransitUrl = BaseUrl + "ups.app/xml/TimeInTransit"; 
    String TimeInTransitRequest = LoadResourceDataString( UPS_TIME_IN_TRANSIT );             
    String PDataOut;

    TReplaceFlags rf;
    rf << rfReplaceAll;
    
    TimeInTransitRequest = StringReplace( TimeInTransitRequest, "{shipfromstreetname}", PFromAddress.Street1, rf );
    TimeInTransitRequest = StringReplace( TimeInTransitRequest, "{shipfromzip}", PFromAddress.Zip, rf );
    TimeInTransitRequest = StringReplace( TimeInTransitRequest, "{shipfromcountry}", GetCountryCodeFromName( PFromAddress.Country ), rf );

    TimeInTransitRequest = StringReplace( TimeInTransitRequest, "{shiptostreetname}", PToAddress.Street1, rf );
    TimeInTransitRequest = StringReplace( TimeInTransitRequest, "{shiptozip}", PToAddress.Zip, rf );
    TimeInTransitRequest = StringReplace( TimeInTransitRequest, "{shiptocountry}", GetCountryCodeFromName( PToAddress.Country ), rf );

    TDateTime today = Date();
    TimeInTransitRequest  = StringReplace( TimeInTransitRequest, "{pickupdate}", today.FormatString("yyyymmdd"), rf );

    LOG( "sending TimeInTransit request to UPS" );

    TWinInet InetIntf;
    PErrorMessage = "Unknown Error";    
    _di_IXMLDocument xml;
    _di_IXMLNode TempNode;

    try {
        PDataOut = InetIntf.Post( TimeInTransitUrl, AccessXML + TimeInTransitRequest );
    }

    catch( EPropertyConvertError const &E ) {
        LOGEXCEPTION( E );
        PErrorMessage = E.Message;
        //	throw EASAPException( E.Message );
    }

    catch( ERemotableException const &E ) {
        LOGEXCEPTION( E );
        PErrorMessage = E.Message;
        //	throw EASAPException( E.Message );
    }
    catch( Exception const &E ) {
        LOGEXCEPTION(E);
        MsgBox( "Failed to send rate request to UPS.", MB_ICONERROR | MB_OK );
        PErrorMessage = E.Message;
        return false;
    }

    LOG( "received TimeInTransit response" );

    try {
        xml = NewXMLDocument();
        xml->LoadFromXML( PDataOut );
        xml->SaveToFile("ups_soap_timeintransit_response.xml");
        _di_IXMLNode ResponseNode = xml->DocumentElement->ChildNodes->FindNode( "Response" );
        if( !ResponseNode )
            throw Exception( "can't find Response node." );

        TempNode = ResponseNode->ChildNodes->FindNode( "ResponseStatusCode" );
        if( !TempNode )
            throw Exception( "can't find ResponseStatusCode node." );

        int StatusCode = StrToInt( TempNode->Text );

        TempNode = ResponseNode->ChildNodes->FindNode( "ResponseStatusDescription" );
        if( !TempNode )
            throw Exception( "can't find ResponseStatusDescription node." );

        String StatusDesc = TempNode->Text;

        if( StatusCode != 1 )
        {
            LOGX( "UPS returned error (%d): %s", ( StatusCode, StatusDesc ) );
            TempNode = ResponseNode->ChildNodes->FindNode( "Error" );
            if( !TempNode )
                throw Exception( "can't find Error node." );
            TempNode = TempNode->ChildNodes->FindNode( "ErrorDescription" );
            if( !TempNode )
                throw Exception( "can't find ErrorDescription node." );
            PErrorMessage = TempNode->Text;
            xml->SaveToFile("ups_soap_bad_response_timeintransit.xml");
            return false;
        }

        _di_IXMLNode TransitResponseNode = xml->DocumentElement->ChildNodes->FindNode( "TransitResponse" );
        if( !TransitResponseNode )
            throw Exception( "can't find TransitResponseNode node." );

        _di_IXMLNode ServiceSummaryNode = TransitResponseNode->ChildNodes->FindNode( "ServiceSummary" );
        if( !ServiceSummaryNode )
            throw Exception( "can't find ServiceSummaryNode node." );

        do
        {
            for (int i = 0; i < ServiceSummaryNode->ChildNodes->Count; ++i) {
                _di_IXMLNode ServiceNode = ServiceSummaryNode->ChildNodes->Nodes[i];
                if (ServiceNode->NodeName == "Service") {
                   _di_IXMLNode CodeNode = ServiceNode->ChildNodes->FindNode( "Code" );

                   if( !CodeNode )
                    continue;

                   if (CodeNode->Text == PServiceType.TimeInTransitCode) {
                    _di_IXMLNode EstimatedArrivalNode =  ServiceSummaryNode->ChildNodes->FindNode( "EstimatedArrival" );
                    if( !EstimatedArrivalNode )
                        continue;

                    _di_IXMLNode BusinessTransitDaysNode =  EstimatedArrivalNode->ChildNodes->FindNode( "BusinessTransitDays" );
                    if( !BusinessTransitDaysNode )
                        continue;

                        PTimeCommit = BusinessTransitDaysNode->Text + " day(s).";
                        PErrorMessage = "";
                        return true;
                   }
                   else
                   {
                        continue;
                   }
                }
            }
            ServiceSummaryNode = ServiceSummaryNode->NextSibling();
        }while (ServiceSummaryNode);       
   }
   catch( Exception const &E ) {
        LOGX( "UPS xml invalid, %s", ( E.Message ) );
        MsgBox( "The XML response from UPS was invalid.", MB_ICONERROR | MB_OK );
        return false;
    }

    PErrorMessage = "";
    return true;
}

//Return TimeInTransitCode given Rate code
String  TUPSInterface::GetTimeInTransitCode(String RateCode)
{
    std::map<String, TUPSServiceCode>::iterator it = MyOBServiceMap.begin();
    while( it != MyOBServiceMap.end() )
    {
        if( it->second.RateCode == RateCode )
        {
          return it->second.TimeInTransitCode;
        }
        it++;
    }

    return "";
}

bool TUPSInterface::GetAddressType(String const& BaseUrl, String const& AccessXML, TAddress const &Address, String &PErrorMessage, String& AddressType)
{
    String AddressValidatorUrl = BaseUrl + "ups.app/xml/XAV";
    String TAddressValidationRequest = LoadResourceDataString( UPS_ADDRESS_VALIDATION );             
    String PDataOut;
    AddressType =  "Unknown";
    
    TReplaceFlags rf;
    rf << rfReplaceAll;

    LOG( "Sending Address validate request to UPS" );

    TWinInet InetIntf;
    PErrorMessage = "Unknown Error";    
    _di_IXMLDocument xml;
    _di_IXMLNode TempNode;

    String countryCode = GetCountryCodeFromName( Address.Country );
    String stateCode = GetStateCodeFromName( Address.State, countryCode );
    TAddressValidationRequest = StringReplace( TAddressValidationRequest, "{AddressLine}", Address.Street1, rf );
    TAddressValidationRequest = StringReplace( TAddressValidationRequest, "{City}", Address.City, rf );
    TAddressValidationRequest = StringReplace( TAddressValidationRequest, "{StateCode}", stateCode, rf );
    TAddressValidationRequest = StringReplace( TAddressValidationRequest, "{ZipCode}", Address.Zip, rf );
    TAddressValidationRequest = StringReplace( TAddressValidationRequest, "{CountryCode}", countryCode, rf );
    
    try {
        PDataOut = InetIntf.Post( AddressValidatorUrl, AccessXML + TAddressValidationRequest );
    }
    catch( EPropertyConvertError const &E ) {
        LOGEXCEPTION( E );
        PErrorMessage = E.Message;
        //	throw EASAPException( E.Message );
    }
    catch( ERemotableException const &E ) {
        LOGEXCEPTION( E );
        PErrorMessage = E.Message;
        //	throw EASAPException( E.Message );
    }
    catch( Exception const &E ) {
        LOGEXCEPTION(E);
        //MsgBox("Failed to send rate request to UPS.", MB_ICONERROR | MB_OK );
        PErrorMessage = E.Message;
        return false;
    }

    LOG( "Received Address validate response" );

    try {
        xml = NewXMLDocument();
        xml->LoadFromXML( PDataOut );
        xml->SaveToFile("ups_soap_AddressValidator_response.xml");
        _di_IXMLNode ResponseNode = xml->DocumentElement->ChildNodes->FindNode( "Response" );
        if( !ResponseNode )
            throw Exception( "can't find Response node." );

        TempNode = ResponseNode->ChildNodes->FindNode( "ResponseStatusCode" );
        if( !TempNode )
            throw Exception( "can't find ResponseStatusCode node." );

        int StatusCode = StrToInt( TempNode->Text );
        TempNode = ResponseNode->ChildNodes->FindNode( "ResponseStatusDescription" );
        if( !TempNode )
            throw Exception( "can't find ResponseStatusDescription node." );

        String StatusDesc = TempNode->Text;
        if( StatusCode != 1 )
        {
            LOGX( "UPS returned error (%d): %s", ( StatusCode, StatusDesc ) );
            TempNode = ResponseNode->ChildNodes->FindNode( "Error" );
            if( !TempNode )
                throw Exception( "can't find Error node." );

            TempNode = TempNode->ChildNodes->FindNode( "ErrorDescription" );
            if( !TempNode )
                throw Exception( "can't find ErrorDescription node." );
        
            PErrorMessage = TempNode->Text;
            xml->SaveToFile("ups_soap_bad_response_AddressValidator.xml");

            return false;
        }

        _di_IXMLNode AddressClassificationNode = xml->DocumentElement->ChildNodes->FindNode( "AddressClassification" );
         if( !AddressClassificationNode )
            throw Exception( "can't find Address Validator node." );

        if (AddressClassificationNode->ChildNodes->Count == 2) {
            _di_IXMLNode AddressTypeNode = AddressClassificationNode->ChildNodes->Nodes[1];
            AddressType =  AddressTypeNode->Text;
        }
    }
    catch( Exception const &E ) {
        LOGX( "UPS xml invalid, %s", ( E.Message ) );
        MsgBox( "The XML response from UPS was invalid.", MB_ICONERROR | MB_OK );
        return false;
    }

    PErrorMessage = "";
    return true;
}
//---------------------------------------------------------------------------

