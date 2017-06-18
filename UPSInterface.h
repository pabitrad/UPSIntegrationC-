//---------------------------------------------------------------------------
#ifndef UPSInterfaceH
#define UPSInterfaceH
//---------------------------------------------------------------------------
#include <map>
#include "ShippingHelpers.h"
#include "NativeXml.hpp"
#include "Address.h"
#include <IniFiles.hpp>
#include <boost/shared_ptr.hpp>
#include "Functions.h"
//---------------------------------------------------------------------------
struct TUPSServiceCode
{
  bool ShipSaturday;
  AnsiString RateCode;
  AnsiString TimeInTransitCode;
  bool International;
  TUPSServiceCode()
    : ShipSaturday( false )
  {
  }

  TUPSServiceCode( const String& PRateCode, const String& PTransitCode, bool PInternational = false, bool const PShipSaturday = false )
    : RateCode( PRateCode ), TimeInTransitCode(PTransitCode), International( PInternational ), ShipSaturday( PShipSaturday )
  {
  }
};
//---------------------------------------------------------------------------
class TUPSInterface
{
private:
  std::map<String, TUPSServiceCode> MyOBServiceMap;
  boost::shared_ptr<TMemIniFile> IniFile;
  HWND FParentWindow;
  int MsgBox( String Msg, int Btns = MB_OK | MB_ICONSTOP )
  {
    return MessageBox( ParentWindow, Msg.c_str(), "ASAPCentral", Btns );
  }

  void InitMyOBServiceMap();
  void __fastcall SoapBeforeExecute(const String MethodName, WideString &SOAPRequest);
  void __fastcall SoapAfterExecute(const String MethodName, Classes::TStream* SOAPResponse);
  bool TUPSInterface::RetrieveCredentials( String &URL, String &Licence,
    String &UserName, String &Password );
  bool UPSGetRate( TUPSServiceCode &PServiceType, double const PWeightPounds,
    double const PWeightOunces, double const PkgLength, double const PkgWidth, double const PkgHeight,
    TAddress const &PFromAddress, TAddress const &PToAddress, String &PEstimatedCost, String &PTimeCommit,
    String &AddressType, String &PErrorMessage );
  bool UPSGetRate( TUPSServiceCode &PServiceType, double const PWeightPounds,
    double const PWeightOunces, double const PkgLength, double const PkgWidth, double const PkgHeight,
    TAddress const &PFromAddress, TAddress const &PToAddress, TXmlNode *PRowNode );

    bool UPSGetDeliveryTime (TUPSServiceCode const &PServiceType, String const& BaseUrl, String const& AccessXML, TAddress const &PFromAddress, TAddress const &PToAddress, 
                         String &PErrorMessage, String &PTimeCommit);
    String GetTimeInTransitCode(String RateCode);
    bool GetAddressType(String const& BaseUrl, String const& AccessXML, TAddress const &PFromAddress, String &PErrorMessage, String& AddressType);

public:
  TUPSInterface()
    : ParentWindow( NULL )
  {
    InitMyOBServiceMap();
    #ifdef ASAP_CLIENT
    IniFile.reset( new TMemIniFile( GetTWUserDir() + "\\TaskWright.cfg" ) );
    #else
    IniFile.reset( new TMemIniFile( GetTWUserDir() + "\\TWServer.cfg" ) );
    #endif
  }

  bool MyOBServiceToUPSServiceCode( const String& PMyOBService, TUPSServiceCode &PServiceCode );
  void CalculateUPS( double PoundWeight, double OunceWeight, double PkgLength,
    double PkgWidth, double PkgHeight, TAddress FromAddr, TAddress ToAddr, TXmlNode *PListNode, String PServiceCode = "-1" );
  __property HWND ParentWindow = { read = FParentWindow, write = FParentWindow };
};
//---------------------------------------------------------------------------
extern TUPSInterface UPSInterface;
//---------------------------------------------------------------------------
#endif
