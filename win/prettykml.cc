// Copyright 2008, Google Inc. All rights reserved.
//
// Redistribution and use in source and binary forms, with or without 
// modification, are permitted provided that the following conditions are met:
//
//  1. Redistributions of source code must retain the above copyright notice, 
//     this list of conditions and the following disclaimer.
//  2. Redistributions in binary form must reproduce the above copyright notice,
//     this list of conditions and the following disclaimer in the documentation
//     and/or other materials provided with the distribution.
//  3. Neither the name of Google Inc. nor the names of its contributors may be
//     used to endorse or promote products derived from this software without
//     specific prior written permission.
//
// THIS SOFTWARE IS PROVIDED BY THE AUTHOR ``AS IS'' AND ANY EXPRESS OR IMPLIED
// WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF 
// MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO
// EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, 
// SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
// PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;
// OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
// WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR 
// OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
// ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

// This sample program reads and parses one KML or KMZ file and emits the
// serialization to standard output.  Since all recognized elements
// are saved in the KML DOM with white space discarded and Serialize()
// formats nicely this is overall a pretty printer.

#include <iostream>
#include <string>
#include "boost/scoped_ptr.hpp"
#include "kml/dom.h"
#include "kml/engine.h"
#include "kml/base/file.h"
#include "kml/base/vec3.h"
#include "excelreader.h"
#include "UtfConverter.h"
#include <direct.h>
#define GetCurrentDir _getcwd

using kmldom::ElementPtr;
using kmlengine::KmlFile;
using kmlengine::KmlFilePtr;
using kmldom::PlacemarkPtr;
using kmlbase::Vec3;
using kmldom::PointPtr;
using std::cout;
using std::endl;

typedef std::vector<ElementPtr> element_vector_t;

template<typename T>
inline std::wstring ToWstring(T value) {
  std::wstringstream ss;
  ss.precision(15);
  ss << value;
  return ss.str();
}


int main(int argc, char** argv) {

   char cCurrentpath[FILENAME_MAX];
         
   if (!GetCurrentDir(cCurrentpath, sizeof(cCurrentpath))) {
     return 1;
   }
   
   strcat_s(cCurrentpath,"\\");

  if (argc < 2) {
   strcat_s(cCurrentpath, "SiteInfoFile.kml");
  } else if (argc == 2) {
   strcat_s(cCurrentpath, argv[1]);
  } else {
	 cout << argv[0] << " kmlfile" << endl;
  }

  msexcel::excelreader eo;

  //eo.SetData(L"D8:D8", L"5");
  //eo.SetData(L"E3:G5", L"kkj");
  std::list<std::wstring> line;
  // Read the file content.
  std::string file_data;
  if (!kmlbase::File::ReadFileToString(cCurrentpath, &file_data)) {
    cout << argv[1] << " read failed" << endl;
    return 1;
  }

  // Parse it.
  std::string errors;
  KmlFilePtr kml_file = KmlFile::CreateFromParse(file_data, &errors);
  if (!kml_file) {
    cout << errors;
    return 1;
  }

  //kmldom::ElementPtr element;
  element_vector_t placemarks;
  kmlengine::GetElementsById(kml_file->get_root(), kmldom::Type_Placemark, &placemarks);

  assert(!placemarks.empty());
  element_vector_t::iterator it = placemarks.begin();
  element_vector_t::iterator end = placemarks.end();
 
  line.push_back(L"标题标注");
  line.push_back(L"经度");
  line.push_back(L"纬度");
  line.push_back(L"说明");

  eo.SetData(1, 1, line);
  line.clear();
  for (int i = 2; it != end; ++it, ++i) {
	  const PlacemarkPtr placemark = AsPlacemark(*it);
	  if (placemark->has_name()) {

		 //line.push_back(StringToWString(placemark->get_name()));
		 // line.push_back(ctow(placemark->get_name().c_str()))
		 //cout << placemark->get_name().length() << endl;
		 
		  //printf("%x\n", placemark->get_name().c_str());
		  line.push_back(UtfConverter::FromUtf8(placemark->get_name()));
		  if (placemark->has_geometry()) {
		    const PointPtr point = AsPoint(placemark->get_geometry());
		  
		    Vec3 vec3 = point->get_coordinates()->get_coordinates_array_at(0);
            //cout << vec3.get_longitude() << " " << vec3.get_latitude() << endl;
		    line.push_back(ToWstring(vec3.get_longitude()));
		    line.push_back(ToWstring(vec3.get_latitude()));
		  }

		  if (placemark->has_description())
			  line.push_back(UtfConverter::FromUtf8(placemark->get_description()));
		  eo.SetData(i, 1, line);
		  line.clear();
	  }
  }


  eo.SaveAsFile(L"BasicReport");

  // Serialize it and output to stdout.
 // std::string output;
  //kml_file->SerializeToString(&output);
  //cout << output;

  return 0;
}
