// ParseOffice365XML.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "string"
#include "windows.h"
//#include "msxml6.h"
#include "fstream"
#include "iostream"
#include "vector"
#include "map"
#include <atlbase.h>
#include <msxml2.h>
#import <msxml6.dll>



typedef std::wstring RCString;
typedef std::vector<WCHAR> StringVector;

RCString GetEncodedXML()
{
    std::wifstream f("C:\\foo.txt");
    RCString line = L"";
    RCString contents;
    while(f.good())
    {
        std::getline(f, line);
        contents += line;
    }

    //std::wcout << contents << std::endl; 
    f.close();
    return contents;
}

typedef std::map<RCString, wchar_t> EscapeCharsToXMLMap;

class EscapeCharactersToXML
{
public:
    static EscapeCharsToXMLMap& GetMap()
    {
        if(map_.empty())
        {
            map_[L"lt"]   = L'<';
            map_[L"gt"]   = L'>';
            map_[L"amp"]  = L'&';
            map_[L"quot"] = L'"';
            map_[L"apos"] = L'\'';

        }
        return map_;
    }
private:
    static EscapeCharsToXMLMap map_;
};

EscapeCharsToXMLMap EscapeCharactersToXML::map_;

static const RCString BEGIN_MARKER = L"BEGINCONFIG";
static const RCString END_MARKER = L"ENDCONFIG";
class DecodedXMLString
{
public:


    DecodedXMLString(RCString& htmlString) : htmlString_(htmlString)
    {
        
    }


    RCString GetDecodedString();
private:


    bool StripExtraneousCharacters();

    RCString htmlString_;

};


bool DecodedXMLString::StripExtraneousCharacters()
{
    const wchar_t* begin = wcsstr(htmlString_.c_str(), BEGIN_MARKER.c_str());
    const wchar_t* end = wcsstr(htmlString_.c_str(), END_MARKER.c_str());
    if(begin != NULL && end != NULL  && begin < end)
    {
        //remove the matching tags from the string pointers
        begin += BEGIN_MARKER.size();
        StringVector newStringVec;
        int size = (end - begin);
        newStringVec.resize(size + 1);
        wchar_t* newString = &newStringVec[0];
        wcsncpy_s(newString, size + 1, begin, size);
        newString[size] = L'\0';
        htmlString_ = newString;
        std::wcout << htmlString_.c_str() << std::endl; 
        return true;
    }
    return false;
}



RCString DecodedXMLString::GetDecodedString()
{
    //TODO: check return
    StripExtraneousCharacters();
    
    StringVector tempBuff(8);  //temporary buffer to hold special escape chars

    int htmlStrLength = htmlString_.length();

    StringVector newString(htmlStrLength + 1); //the new string will be at most as big as the html string read

    int newStringWritePos = 0;

    int i = 0;
    while(i < htmlStrLength)
    {
        WCHAR currentChar = htmlString_[i++];
        if(currentChar == L'&')
        {
            LPWSTR escapeCharEnd = wcsstr(&htmlString_[i], L";");
            if(escapeCharEnd != NULL)
            {
                unsigned int escapeCharSize = escapeCharEnd - &htmlString_[i]; //the character we read will be between the current position of htmlString + 1, and escapeCharEnd;
                if(tempBuff.size() < (escapeCharSize + 1))
                    tempBuff.resize(escapeCharSize + 1);
                wcsncpy_s(&tempBuff[0], escapeCharSize + 1, &htmlString_[i] , escapeCharSize);
                tempBuff[escapeCharSize] = L'\0';
                int intValue = 0;
                if(escapeCharSize > 1 && tempBuff[0] == L'#') //we've found a hex or decimal number
                {
                    if(escapeCharSize > 2 && (tempBuff[1] == L'X' || tempBuff[1] == L'x')) //found a hex number
                    {
                        RCString hexString = L"0x";
                        hexString += &tempBuff[2]; //append hex number
                        intValue = wcstol(hexString.c_str(), NULL, 16);
                    }
                    else
                        intValue = wcstol(&tempBuff[1], NULL, 10);

                    if(intValue != 0)
                    {
                        currentChar = (WCHAR) intValue;     
                        i = escapeCharEnd - &htmlString_[0] + 1; //pointer arithmetic to find out how far we are from the beginning of the string
                    }
                }
                else  //we found an escape sequence for a special char
                {
                    i = escapeCharEnd - &htmlString_[0] + 1;
                    EscapeCharsToXMLMap::iterator iter = EscapeCharactersToXML::GetMap().begin();
                    EscapeCharsToXMLMap::iterator iterEnd = EscapeCharactersToXML::GetMap().end();
                    WCHAR matchedReplacement = L'\0';

                    for(;iter != iterEnd; iter++)
                    {
                        if(iter->first == &tempBuff[0]) //see if the key matches
                        {
                           matchedReplacement = iter->second; 
                           break;
                        }
                    }

                    if(matchedReplacement != L'\0')
                    {
                        currentChar = matchedReplacement;
                    }
                    else
                    { 
                        //copy the unrecognized escape sequence as is
                        newString[newStringWritePos++] = L'&';
                        wcsncpy_s(&newString[newStringWritePos], escapeCharSize + 1, &tempBuff[0], escapeCharSize);
                        newStringWritePos += escapeCharSize;
                        newString[newStringWritePos++] = L';';
                        continue;
                    }
                }
            }                
        }
        newString[newStringWritePos++] = currentChar;
    }
    newString[newStringWritePos] = L'\0';

    return &newString[0];
}


//class Office365UserSettings
//{
//public:
//    HRESULT Initialize(const RCString& xmlPayload)
//    {
//        CComPtr<IXMLDOMDocument2> domDoc;
//        HRESULT hr = domDoc.CoCreateInstance(CLSID_DOMDocument2);
//		if (SUCCEEDED(hr))
//		{
//			// these methods should not fail so don't inspect result
//			domDoc->put_async(VARIANT_FALSE);
//			domDoc->put_validateOnParse(VARIANT_FALSE);
//			domDoc->put_resolveExternals(VARIANT_FALSE);
//		    VARIANT var;
//
//		    VariantInit(&var);
//		    hr = VariantFromString(L"SelectionLanguage", L"XPath");
//		if (SUCCEEDED(hr))
//		{
//			CComBSTR bstrProperty(prop.c_str());
//
//			IFFALSE_EXIT_HR(xmlDom, E_FAIL);
//			hr = xmlDom->setProperty(bstrProperty, var);
//		}
//
//	Exit:
//		VariantClear(&var);
//
//
//		}
//        else
//            return E_FAIL;
//
//        return S_OK;
//
//    }
//
//
//
//};



int _tmain(int argc, _TCHAR* argv[])
{
    ::CoInitialize(NULL);

    RCString str = GetEncodedXML();
    DecodedXMLString decoder(str);








    RCString xml = decoder.GetDecodedString();
    wprintf(xml.c_str());
    ::CoUninitialize();
    return 0;
}

