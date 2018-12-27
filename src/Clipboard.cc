#include "Clipboard.h"
#include <comdef.h>
#include <windows.h>
#include <wincodec.h>
#include <string>
#include <sstream>
#include <iostream>
#include <iomanip>

using namespace v8;

Nan::Persistent<v8::FunctionTemplate> Clipboard::constructor;

NAN_MODULE_INIT(Clipboard::Init) {
  v8::Local<v8::FunctionTemplate> ctor = Nan::New<v8::FunctionTemplate>(Clipboard::New);
  // TODO: I don't understand what this does
  constructor.Reset(ctor);
  ctor->InstanceTemplate()->SetInternalFieldCount(1);
  ctor->SetClassName(Nan::New("Clipboard").ToLocalChecked());

  Nan::SetPrototypeMethod(ctor, "writeBitmapToDisk", WriteBitmapToDisk);

  // TODO: what does this do?
  target->Set(Nan::New("Clipboard").ToLocalChecked(), ctor->GetFunction());
}

// TODO: how create a class that contains only static methods?
NAN_METHOD(Clipboard::New) {

  // throw an error if constructor is called without new keyword
  if(!info.IsConstructCall()) {
    return Nan::ThrowError(Nan::New("Clipboard::New - called without new keyword").ToLocalChecked());
  }

  // create a new instance and wrap our javascript instance
  Clipboard* vec = new Clipboard();
  vec->Wrap(info.Holder());

  // return the wrapped javascript instance
  info.GetReturnValue().Set(info.Holder());
}

std::wstring s2ws(const std::string& s)
{
    int len;
    int slength = (int)s.length() + 1;
    len = MultiByteToWideChar(CP_ACP, 0, s.c_str(), slength, 0, 0); 
    wchar_t* buf = new wchar_t[len];
    MultiByteToWideChar(CP_ACP, 0, s.c_str(), slength, buf, len);
    std::wstring r(buf);
    delete[] buf;
    return r;
}

void RaiseError(std::string errorMessage, HRESULT hr)
{
  std::stringstream ss;
  ss << errorMessage << "0x" << std::uppercase << std::setfill('0') << std::setw(8) << std::hex << hr;
  Nan::ThrowError(Nan::New(ss.str()).ToLocalChecked());
}

void InternalWriteBitmapToDisk(std::wstring filename, std::string file_format, int width_constraint, bool write_full)
{
  if (!OpenClipboard(NULL))
  {
    return;
  }

  HANDLE hBitmap = GetClipboardData(CF_BITMAP);
  if (hBitmap == NULL)
  {
    return;
  }

  HRESULT hr = CoInitialize(NULL);
  hr = E_FAIL;
  if (FAILED(hr))
  {
    return RaiseError("Failed to initialize COM: ", hr);
  }

  IWICImagingFactory *ipFactory = NULL;
  hr = CoCreateInstance(CLSID_WICImagingFactory,
                        NULL,
                        CLSCTX_INPROC_SERVER,
                        IID_IWICImagingFactory,
                        reinterpret_cast<void **>(&ipFactory));
  if (FAILED(hr))
  {
    return RaiseError("Failed to initialize WIC Imaging Factory object: ", hr);
  }

  IWICBitmap *ipBitmap = NULL;
  hr = ipFactory->CreateBitmapFromHBITMAP(reinterpret_cast<HBITMAP>(hBitmap),
                                          NULL,
                                          WICBitmapIgnoreAlpha,
                                          &ipBitmap);
  if (FAILED(hr))
  {
    return RaiseError("Failed to construct a WIC Bitmap object from the HBITMAP from the clipboard: ", hr);
  }

  // Retrieve the width and height of the clipboard image
  UINT width, height;
  hr = ipBitmap->GetSize(&width, &height);
  if (FAILED(hr))
  {
    return RaiseError("Could not get the width and height of the WIC Bitmap object: ", hr);
  }

  // Compute the output image width and height by constraining
  // the maximum width of the image to 800px.
  UINT output_width, output_height;
  float scaling_factor = (float)((float)width_constraint / (float)width);
  output_width = width_constraint;
  output_height = scaling_factor * height;

  // Now resize it using a WIC Bitmap Scaler object
  IWICBitmapScaler *ipScaler = NULL;
  hr = ipFactory->CreateBitmapScaler(&ipScaler);
  if (FAILED(hr)) 
  {
    return RaiseError("Could not create a WIC Bitmap scaler object: ", hr);
  }

  hr = ipScaler->Initialize(ipBitmap,
                            output_width,
                            output_height,
                            WICBitmapInterpolationModeHighQualityCubic);
  if (FAILED(hr))
  {
    return RaiseError("Could not initialize the WIC Bitmap scaler object InterpolationMode High Quality Cubic: ", hr);
  }

  IWICBitmapEncoder *ipBitmapEncoder = NULL;
  GUID encoderId = file_format == "png" ? GUID_ContainerFormatPng : GUID_ContainerFormatJpeg;
  hr = ipFactory->CreateEncoder(encoderId,
                                NULL,
                                &ipBitmapEncoder);
  if (FAILED(hr))
  {
    return RaiseError("Could not create the PNG or JPG encoder: ", hr);
  }

  IWICStream *ipStream = NULL;
  hr = ipFactory->CreateStream(&ipStream);
  if (FAILED(hr)) {
    return RaiseError("Could not create an IStream object: ", hr);
  }

  hr = ipStream->InitializeFromFilename((filename + L"." + file_format).c_str(), GENERIC_WRITE);
  if (FAILED(hr))
  {
    return RaiseError("Failed to initialize a writeable stream: ", hr);
  }

  hr = ipBitmapEncoder->Initialize(ipStream, WICBitmapEncoderNoCache);
  if (FAILED(hr))
  {
    return RaiseError("Failed to initialize the bitmap encoder using the stream: ", hr);
  }

  IWICBitmapFrameEncode *ipFrameEncoder = NULL;
  hr = ipBitmapEncoder->CreateNewFrame(&ipFrameEncoder, NULL);
  if (FAILED(hr))
  {
    return RaiseError("Failed to create a new frame encoder using the bitmap encoder: ", hr);

  }

  hr = ipFrameEncoder->Initialize(NULL);
  if (FAILED(hr))
  {
    return RaiseError("Failed to initialize the frame encoder: ", hr);
  }

  if (SUCCEEDED(hr))
  {
    std::cout << "Initialized the frame encoder\n";
    hr = ipFrameEncoder->SetSize(output_width, output_height);
    if (SUCCEEDED(hr))
    {
      std::cout << "Set output size";
      WICPixelFormatGUID formatGuid = GUID_WICPixelFormat24bppRGB;
      hr = ipFrameEncoder->SetPixelFormat(&formatGuid);
      if (SUCCEEDED(hr))
      {
        std::cout << "Set pixel format to RGB 8 bits per channel\n";
        hr = ipFrameEncoder->WriteSource(ipScaler, NULL);
        if (SUCCEEDED(hr))
        {
          std::cout << "Set write source to the IWICBitamp object\n";
          hr = ipFrameEncoder->Commit();
          if (SUCCEEDED(hr))
          {
            std::cout << "Committed the frame encoder\n";
            hr = ipBitmapEncoder->Commit();
            if (SUCCEEDED(hr))
            {
              std::cout << "Committed the bitmap encoder\n";
            }
          }
        }
      }
    }
    }
      }
    }
    }
    }
    }
    }
  }
}
ipFactory->Release();
CoUninitialize();
return;
}

const std::string writeBitmapToDiskParameterError{"writeBitmapToDisk - expected arguments filename, file_format, width_constraint, write_full"};

NAN_METHOD(Clipboard::WriteBitmapToDisk)
{
  // named? (TODO: make named via struct in the future) parameters:
  // filename: filename (without extension) of image
  // file_format: png|jpg|auto - selects the right serializer
  // width_constraint: integer - when specified, it will 
  // write_full: true|false - writes full sized image to disk with _f appended to filename

  // Validate number of parameters
  if (info.Length() != 4)
  {
    return Nan::ThrowError(Nan::New(writeBitmapToDiskParameterError).ToLocalChecked());
  }

  // Validate parameter types
  if (!info[0]->IsString() || !info[1]->IsString() || !info[2]->IsNumber() || !info[3]->IsBoolean())
  {
    return Nan::ThrowError(Nan::New(writeBitmapToDiskParameterError).ToLocalChecked());
  }

  // Validate parameter values
  std::wstring filename = s2ws(*v8::String::Utf8Value(Isolate::GetCurrent(), info[0]->ToString()));
  std::string file_format = *v8::String::Utf8Value(Isolate::GetCurrent(), info[1]->ToString());
  // TODO: figure out how to deal with deprecated Int32Value method
  int width_constraint = info[2]->Int32Value();
  bool write_full = info[3]->BooleanValue();

  if (!(file_format == "png" || file_format == "jpg"))
  {
    return Nan::ThrowError(Nan::New("writeBitmapToDisk - file_format must be png|jpg").ToLocalChecked());
  }

  InternalWriteBitmapToDisk(filename, file_format, width_constraint, write_full);
}