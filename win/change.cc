#define GSM_EXCEL_FILE 1

#ifdef GSM_EXCEL_FILE
#define LATITUDE_COLUMN 25
#define LONGITUDE_COLUMN 24
#define DOA_COLUMN 32
#define CELLNAME_COLUMN 10
#define CELLINDENTITY_COLUMN 8
#define HEIGHT_COLUMN 37
#define DOWNTILE_COLUMN 33
#define INSTALLMODE_COLUMN 36
#define SHARETYPE_COLUMN 12
#define INPUT_FILE "中山基站物理信息.xls"
#define OUTPUT_FILE "GSMCells.kml"
#else
#define LATITUDE_COLUMN 15
#define LONGITUDE_COLUMN 14
#define DOA_COLUMN 16
#define CELLNAME_COLUMN 12
#define CELLINDENTITY_COLUMN 9
#define HEIGHT_COLUMN 21
#define DOWNTILE_COLUMN 17
#define INSTALLMODE_COLUMN 20
#define SHARETYPE_COLUMN 30
#define INPUT_FILE "中山WCDMA基站物理信息.xls"
#define OUTPUT_FILE "WCDMACells.kml"
#endif

#include <iostream>
#include "kml/dom.h"
#include "kml/convenience/convenience.h"
#include "kml/base/file.h"
#include "kml/base/vec3.h"
#include "kml/base/math_util.h"
#include <fstream>
#include "excelreader.h"
#include <direct.h>
#define GetCurrentDir _getcwd
//#include<ctime>
// begin
// engine links libexpat.dll, we could skip it if we don't use
// using kmlengine::KmlFile;
// using kmlengine::KmlFilePtr;
//#include "kml/engine.h"
// end
using kmldom::DocumentPtr;
using kmldom::IconStylePtr;
using kmldom::KmlFactory;
using kmldom::KmlPtr;
using kmldom::PairPtr;
using kmldom::PlacemarkPtr;
using kmldom::StylePtr;
using kmldom::StyleMapPtr;
using kmldom::LineStylePtr;
using kmldom::PolyStylePtr;
using kmldom::LabelStylePtr;
using kmldom::FolderPtr;
using kmldom::CoordinatesPtr;
using kmldom::PointPtr;
using kmldom::IconStyleIconPtr;
using kmldom::PolygonPtr;
using kmldom::OuterBoundaryIsPtr;
using kmldom::LinearRingPtr;

//
//using kmlengine::KmlFile;
//using kmlengine::KmlFilePtr;

const int sectorSize = 30; // sector size in degree.
const int altitudeSize = 50;

typedef struct cellInfomation {
	std::string cellName;
	std::string  cellIndentity;
	std::string  height;
	std::string  downTile;
	std::string InstallMode;
	std::string shareType;
//	std::string address;
} cellInfomation;


PlacemarkPtr CreatePlacemark(kmldom::KmlFactory* factory,
                             const std::string& name,
                             double lat, double lng) {
  PlacemarkPtr placemark(factory->CreatePlacemark());
  placemark->set_name(name);
  placemark->set_styleurl("#onlytextname");
  CoordinatesPtr coordinates(factory->CreateCoordinates());
  coordinates->add_latlng(lat, lng);

  PointPtr point(factory->CreatePoint());
  point->set_extrude(true);
  point->set_altitudemode(kmldom::ALTITUDEMODE_RELATIVETOGROUND);
  point->set_coordinates(coordinates);
  
  placemark->set_geometry(point);

  return placemark;
}

// Corners presumed to be the corners of the ring.
template<int N>
static LinearRingPtr CreateBoundary(const double (&corners)[N]) {
  KmlFactory* factory = KmlFactory::GetFactory();
  CoordinatesPtr coordinates = factory->CreateCoordinates();
  for (int i = 0; i < N; i+=2) {
    coordinates->add_latlng(corners[i], corners[i+1]);
  }
  // Last must be the same as first in a LinearRing.
  coordinates->add_latlng(corners[0], corners[1]);

  LinearRingPtr linearring = factory->CreateLinearRing();
  linearring->set_coordinates(coordinates);
  return linearring;
}

CoordinatesPtr CreateCoordinatesSector(double lat, double lng,
                                       double radius, int orien) {
  CoordinatesPtr coords = KmlFactory::GetFactory()->CreateCoordinates();
  coords->add_latlngalt(lat, lng, altitudeSize);
  size_t l = orien - sectorSize > 0 ? orien - sectorSize : 360 + (orien - sectorSize);
  size_t h = l + 2 * sectorSize;
  for (size_t i = l; i < h; ++i) {
	kmlbase::Vec3 v = kmlbase::LatLngOnRadialFromPoint(lat, lng, radius, i);
	v.set(2, altitudeSize);
    coords->add_vec3(v);
  }
  coords->add_latlngalt(lat, lng, altitudeSize);
  return coords;
}

//static const size_t kSegments = 350;

// Creates a LinearRing that is the great circle described by a constant
// radius from lat, lng.
static LinearRingPtr SectorFromPointAndRadius(double lat, double lng,
                                              double radius, double doa) {
  CoordinatesPtr coords = CreateCoordinatesSector(
      lat, lng, radius, (int)doa);                                        
  LinearRingPtr lr = KmlFactory::GetFactory()->CreateLinearRing();
  lr->set_coordinates(coords);
  return lr;
}

// This converts to std::string from any T of int, bool or double.
template<typename T>
inline std::string ToString(T value) {
  std::stringstream ss;
  ss.precision(15);
  ss << value;
  return ss.str();
}


static PlacemarkPtr CreateGraphicPlacemark(kmldom::KmlFactory* factory,
                             const std::string& name,
                             double lat, double lng, double doa, cellInfomation &ci) {
  // const double outer_corners[] = {
  // 							    37.80198570954779,-122.4319382787491,
  // 							    37.80166118304026,-122.4318730681021,
  // 							    37.8017138829201,-122.4314979385389,
  // 							    37.80202995372478,-122.4315644851293
  // 							  };
  PlacemarkPtr placemark(factory->CreatePlacemark());
  placemark->set_name(name);
  //const std::string KN2("Install Mode");
  //const std::string KV2(ci.InstallMode);
const std::string Names[] = {"Cell Name", "CI", "高度", "下倾角", "安装方式", "共站类型"};
const std::string Values[] = { ci.cellName, ci.cellIndentity, ci.height, ci.downTile, ci.InstallMode, ci.shareType};

for (int i = 0; i < 6; ++i) 
  kmlconvenience::AddExtendedDataValue(Names[i], Values[i], placemark);
  
//placemark->set_description("hi there!");
  placemark->set_styleurl("#msn-yzl-pushpin-sector00");

  PolygonPtr polygon = factory->CreatePolygon();
  polygon->set_altitudemode(kmldom::ALTITUDEMODE_RELATIVETOGROUND);
  OuterBoundaryIsPtr outerboundaryis = factory->CreateOuterBoundaryIs();
  //outerboundaryis->set_linearring(CreateBoundary(outer_corners));
outerboundaryis->set_linearring(SectorFromPointAndRadius(lat, lng, 100.0, doa));
  polygon->set_outerboundaryis(outerboundaryis);
 
  placemark->set_geometry(polygon);

  return placemark;
}

// Utility functions for encoding Unicode text (wide strings) in
// UTF-8.

// A Unicode code-point can have upto 21 bits, and is encoded in UTF-8
// like this:
//
// Code-point length   Encoding
//   0 -  7 bits       0xxxxxxx
//   8 - 11 bits       110xxxxx 10xxxxxx
//  12 - 16 bits       1110xxxx 10xxxxxx 10xxxxxx
//  17 - 21 bits       11110xxx 10xxxxxx 10xxxxxx 10xxxxxx

// The maximum code-point a one-byte UTF-8 sequence can represent.
const uint32_t kMaxCodePoint1 = (static_cast<uint32_t>(1) <<  7) - 1;

// The maximum code-point a two-byte UTF-8 sequence can represent.
const uint32_t kMaxCodePoint2 = (static_cast<uint32_t>(1) << (5 + 6)) - 1;

// The maximum code-point a three-byte UTF-8 sequence can represent.
const uint32_t kMaxCodePoint3 = (static_cast<uint32_t>(1) << (4 + 2*6)) - 1;

// The maximum code-point a four-byte UTF-8 sequence can represent.
const uint32_t kMaxCodePoint4 = (static_cast<uint32_t>(1) << (3 + 3*6)) - 1;

inline uint32_t ChopLowBits(uint32_t* bits, int n) {
  const uint32_t low_bits = *bits & ((static_cast<uint32_t>(1) << n) - 1);
  *bits >>= n;
  return low_bits;
}

// Converts a Unicode code-point to its UTF-8 encoding.
std::string ToUtf8String(wchar_t wchar) {
  char str[5] = {};  // Initializes str to all '\0' characters.

  uint32_t code = static_cast<uint32_t>(wchar);
  if (code <= kMaxCodePoint1) {
    str[0] = static_cast<char>(code);                          // 0xxxxxxx
  } else if (code <= kMaxCodePoint2) {
    str[1] = static_cast<char>(0x80 | ChopLowBits(&code, 6));  // 10xxxxxx
    str[0] = static_cast<char>(0xC0 | code);                   // 110xxxxx
  } else if (code <= kMaxCodePoint3) {
    str[2] = static_cast<char>(0x80 | ChopLowBits(&code, 6));  // 10xxxxxx
    str[1] = static_cast<char>(0x80 | ChopLowBits(&code, 6));  // 10xxxxxx
    str[0] = static_cast<char>(0xE0 | code);                   // 1110xxxx
  } else if (code <= kMaxCodePoint4) {
    str[3] = static_cast<char>(0x80 | ChopLowBits(&code, 6));  // 10xxxxxx
    str[2] = static_cast<char>(0x80 | ChopLowBits(&code, 6));  // 10xxxxxx
    str[1] = static_cast<char>(0x80 | ChopLowBits(&code, 6));  // 10xxxxxx
    str[0] = static_cast<char>(0xF0 | code);                   // 11110xxx
  } else {
    return std::string("(Invalid Unicode 0x%llX)");//,
                          //static_cast<uint64_t>(wchar));
  }

  return std::string(str);
}


std::string WideToCString(const wchar_t *wide_c_str) {
  if (wide_c_str == NULL) return std::string("(null)");
  std::stringstream ss;
  while (*wide_c_str) {
    ss << ToUtf8String(*wide_c_str++).c_str();
  }
  return ss.str();
}

wchar_t * ANSIToUnicode( const char* str )
{
      int    textlen ;
      wchar_t * result;
      textlen = MultiByteToWideChar( CP_ACP, 0, str,-1,    NULL,0 ); 
      result = (wchar_t *)malloc((textlen+1)*sizeof(wchar_t)); 
      memset(result,0,(textlen+1)*sizeof(wchar_t)); 
      MultiByteToWideChar(CP_ACP, 0,str,-1,(LPWSTR)result,textlen ); 
      return    result; 
}

//bool WriteStringToFile(const std::string& data,
//                             const std::string& filename) {
//  if (filename.empty()) {
//    return false;
//  }
//  wchar_t *ws = ANSIToUnicode(data.c_str());
//  std::string s = WideToCString(ws);
//  FILE *f;
//  fopen_s(&f, filename.c_str(), "w");
//  fwrite(s.c_str(), sizeof(char), s.length(), f);
//  fclose(f);
//  return true;
//}




int main(int argc, char** argv) {

  //time_t begin	= clock();

  char fileName[FILENAME_MAX];
  //wchar_t* excelFileName;// = L"中山WCDMA基站物理信息.xls";
  if (argc < 2) {
	strcpy_s(fileName,INPUT_FILE);
  } else if (argc == 2) {
	strcpy_s(fileName, argv[1]);
  } else {
	std::cout << "usage: " << argv[0] << " filename.xls" <<
    std::endl;
    return 1;
  }

   char cCurrentpath[FILENAME_MAX];
   if (!GetCurrentDir(cCurrentpath, sizeof(cCurrentpath))) {
     return 1;
   }
   strcat_s(cCurrentpath,"\\");
 strcat_s(cCurrentpath, fileName);
    size_t origsize = strlen(cCurrentpath) + 1;
    size_t convertedChars = 0;
    wchar_t* wcstring=ANSIToUnicode(cCurrentpath);
   // mbstowcs_s(&convertedChars, wcstring, origsize, cCurrentpath, _TRUNCATE);
	//wcscat_s(wcstring, L"\\");
//    wcscat_s(wcstring, fileName);
    //wcout << wcstring << endl;

  msexcel::excelreader eo(wcstring, L"基站物理参数表");
	// Load a workbook with one sheet, display its contents and save into another file.

  const int maxRows = eo.rowCount() + 1;

  //std::cout << maxRows << std::endl;
  KmlFactory* kml_factory = KmlFactory::GetFactory();

  DocumentPtr document = kml_factory->CreateDocument();
  document->set_name("zs_rbs_info");
  document->set_open(1);
  
  StylePtr sector0 = kml_factory->CreateStyle();
  sector0->set_id("msn-yzl-pushpin-sector0");
  LineStylePtr linestyle = kml_factory->CreateLineStyle();
  std::string solid_blue("ffff0000");
  std::string tran_green("0000ff00");  // aabbggrr
  linestyle->set_color(solid_blue);
  linestyle->set_width(3.0);
  sector0->set_linestyle(linestyle);
  PolyStylePtr polystyle = kml_factory->CreatePolyStyle();
  polystyle->set_color(tran_green);
  sector0->set_polystyle(polystyle);
  document->add_styleselector(sector0);

  StyleMapPtr stylemap = kml_factory->CreateStyleMap();
  stylemap->set_id("msn-yzl-pushpin-sector00");
  PairPtr pair = kml_factory->CreatePair();
  pair->set_key(kmldom::STYLESTATE_NORMAL);
  pair->set_styleurl("#msn-yzl-pushpin-sector0");
  stylemap->add_pair(pair);

  pair = kml_factory->CreatePair();
  pair->set_key(kmldom::STYLESTATE_HIGHLIGHT);
  pair->set_styleurl("#msn-yzl-pushpin-sector02");
  stylemap->add_pair(pair);

  document->add_styleselector(stylemap);

  
  StylePtr textOnly = kml_factory->CreateStyle();
  IconStyleIconPtr icon = KmlFactory::GetFactory()->CreateIconStyleIcon();
  textOnly->set_id("onlytextname");
  IconStylePtr iconstyle = kml_factory->CreateIconStyle();
  iconstyle->set_icon(icon);
  textOnly->set_iconstyle(iconstyle);
  LabelStylePtr labelstyle = kml_factory->CreateLabelStyle();
  textOnly->set_labelstyle(labelstyle);
  document->add_styleselector(textOnly);

  StylePtr highlight = kml_factory->CreateStyle();
  highlight->set_id("msn-yzl-pushpin-sector02");
  highlight->set_linestyle(linestyle);
  highlight->set_polystyle(polystyle);
  document->add_styleselector(highlight);


  FolderPtr folder = kml_factory->CreateFolder();
  folder->set_name("cell name");

  FolderPtr graphicFolder = kml_factory->CreateFolder();
  graphicFolder->set_name("graphic");

  for(int r = 3; r < maxRows; ++r) {
  	//for(size_t c = 0; c < maxCols; ++c) {
	//BasicExcelCell* cell = sheet1->Cell(r,0);
	  if (eo.GetDataAsString(r, 1).empty()) {
 		  //printf("empty string");
		continue;
	  }
	// todo:remove empty lat and lon site because it's too ugly.
	double lat = eo.GetDataAsDouble(r, LATITUDE_COLUMN); //->Cell(r,14)->GetDouble();
	double lng = eo.GetDataAsDouble(r, LONGITUDE_COLUMN); //sheet1->Cell(r,13)->GetDouble();
	double doa = eo.GetDataAsDouble(r, DOA_COLUMN); //sheet1->Cell(r, 15)->GetDouble();
	cellInfomation ci;
	ci.cellName = eo.GetDataAsString(r, CELLNAME_COLUMN); //sheet1->Cell(r,11)->GetString();
	ci.cellIndentity = eo.GetDataAsString(r, CELLINDENTITY_COLUMN); // sheet1->Cell(r, 8)->GetDouble();
	ci.height = eo.GetDataAsString(r, HEIGHT_COLUMN); //sheet1->Cell(r, 20)->GetDouble();
	ci.downTile = eo.GetDataAsString(r, DOWNTILE_COLUMN);// sheet1->Cell(r, 16)->GetDouble();
	ci.InstallMode = eo.GetDataAsString(r, INSTALLMODE_COLUMN);//WideToCString(sheet1->Cell(r, 19)->GetWString());
	ci.shareType = eo.GetDataAsString(r, SHARETYPE_COLUMN); //WideToCString(sheet1->Cell(r, 29)->GetWString());
		
    folder->add_feature(CreatePlacemark(kml_factory, eo.GetDataAsString(r, 1), lat, lng));

	graphicFolder->add_feature(CreateGraphicPlacemark(kml_factory,eo.GetDataAsString(r, 1), lat, lng, doa, ci));
	
  	//}
  }
 
  document->add_feature(folder);
  document->add_feature(graphicFolder);

  KmlPtr kml = kml_factory->CreateKml();
  kml->set_feature(document);
  const char* output_filename = OUTPUT_FILE;
  
  std::string kml_output = kmldom::SerializePretty(kml);
  wchar_t *ws = ANSIToUnicode(kml_output.c_str());
  std::string kml_stream = WideToCString(ws); // ascii as utf-8
  if (!kmlbase::File::WriteStringToFile(kml_stream, output_filename)) {
  //if (!WriteStringToFile(kml_output, output_filename)) {
    std::cerr << "write failed: " << output_filename << std::cerr;
    return 1;
  }

  //KmlFilePtr kml_file = KmlFile::CreateFromImport(kml);
  //std::string kml_data;
  //kml_file->SerializeToString(&kml_data);
  //if (!myWriteStringToFile(kml_data, output_filename)) {
  //  std::cerr << "write failed: " << output_filename << std::endl;
  //  return false;
  //}
  //std::cout << kmldom::SerializePretty(kml);

 //time_t end	= clock();
 //std::cout << ((double)(end - begin)) / 1000 << std::endl;

  return 0;
}


