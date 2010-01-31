#include <iostream>
#include "kml/dom.h"
#include "kml/convenience/convenience.h"
#include "kml/base/file.h"
#include "kml/base/vec3.h"
#include "kml/base/math_util.h"

#include "BasicExcel.h"
using namespace YExcel;

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

const int sectorSize = 30; // sector size in degree.
const int altitudeSize = 50;

typedef struct cellInfomation {
	std::string cellName;
	double cellIndentity;
	double height;
	double downTile;
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
  const std::string KN2("Install Mode");
  const std::string KV2(ci.InstallMode);
const std::string Names[] = {"Cell Name", "CI", "高度", "下倾角", "安装方式", "共站类型"};
const std::string Values[] = { ci.cellName, ToString(ci.cellIndentity), ToString(ci.height), ToString(ci.downTile), ci.InstallMode, ci.shareType};

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


int main(int argc, char** argv) {
  
  BasicExcel e;

	// Load a workbook with one sheet, display its contents and save into another file.
  e.Load("WC.xls");	
  BasicExcelWorksheet* sheet1 = e.GetWorksheet(L"基站物理参数表");
  if (!sheet1)
	return 1;
  size_t maxRows = sheet1->GetTotalRows();
  //size_t maxCols = sheet1->GetTotalCols();
  //std::cout << maxRows << " " << maxCols << std::endl;
   // BasicExcelCell* cell = sheet1->Cell(2,15); 
   //         switch (cell->Type()) 
   //              { 
   //                  case BasicExcelCell::UNDEFINED: 
   //                    printf(" "); 
   //                    break; 
   //                  case BasicExcelCell::INT: 
   //                    printf("%10d", cell->GetInteger()); 
   //                    break; 
   //                  case BasicExcelCell::DOUBLE: 
   //                    printf("%10.6lf", cell->GetDouble()); 
   //                    break; 
   //                  case BasicExcelCell::STRING: 
   //                    printf("%10s", cell->GetString()); 
   //                    break; 
   //                  case BasicExcelCell::WSTRING: 
   //                    //wprintf(L"%10s", cell->GetWString()); 
   //        		    printf("%10s", WideToCString(cell->GetWString()).c_str());
   //                    break; 
   //                }
	
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

  for(size_t r = 2; r < maxRows; ++r) {
  	//for(size_t c = 0; c < maxCols; ++c) {
	BasicExcelCell* cell = sheet1->Cell(r,0);
	if (cell->GetWString() == NULL) continue;
	double lat = sheet1->Cell(r,14)->GetDouble();
	double lng = sheet1->Cell(r,13)->GetDouble();
	double doa = sheet1->Cell(r, 15)->GetDouble();
	cellInfomation ci;
	ci.cellName = sheet1->Cell(r,11)->GetString();
	ci.cellIndentity = sheet1->Cell(r, 8)->GetDouble();
	ci.height = sheet1->Cell(r, 20)->GetDouble();
	ci.downTile = sheet1->Cell(r, 16)->GetDouble();
	ci.InstallMode = WideToCString(sheet1->Cell(r, 19)->GetWString());
	ci.shareType = WideToCString(sheet1->Cell(r, 29)->GetWString());
		
    folder->add_feature(CreatePlacemark(kml_factory, WideToCString(cell->GetWString()), lat, lng));

	graphicFolder->add_feature(CreateGraphicPlacemark(kml_factory, WideToCString(cell->GetWString()), lat, lng, doa, ci));
	
  	//}
  }
 
  document->add_feature(folder);
  document->add_feature(graphicFolder);

  KmlPtr kml = kml_factory->CreateKml();
  kml->set_feature(document);
  const char* output_filename = "mycells.kml";
  std::string kml_output = kmldom::SerializePretty(kml);
  if (!kmlbase::File::WriteStringToFile(kml_output, output_filename)) {
    std::cerr << "write failed: " << output_filename << std::cerr;
    return 1;
  }

  std::cout << kmldom::SerializePretty(kml);

	
  return 0;
}
