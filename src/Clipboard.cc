#include "Clipboard.h"
#include <comdef.h>
#include <windows.h>
#include <wincodec.h>
#include <string>
#include <iostream>

using namespace v8;

Nan::Persistent<v8::FunctionTemplate> Clipboard::constructor;

NAN_MODULE_INIT(Clipboard::Init) {
  v8::Local<v8::FunctionTemplate> ctor = Nan::New<v8::FunctionTemplate>(Clipboard::New);
  // TODO: I don't understand what this does
  constructor.Reset(ctor);
  ctor->InstanceTemplate()->SetInternalFieldCount(1);
  ctor->SetClassName(Nan::New("Clipboard").ToLocalChecked());

  Nan::SetPrototypeMethod(ctor, "save", Save);

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

NAN_METHOD(Clipboard::Save)
{
  // named? parameters:
  // filename: filename (without extension) of image
  // file_format: png|jpg - selects the right version
  // width_constraint: integer -
  // serialize:

  // SAMPLE CODE to check arguments
  // if(info.Length() != 3) {
  //   return Nan::ThrowError(Nan::New("Vector::New - expected arguments x, y, z").ToLocalChecked());
  // }

  // // expect arguments to be numbers
  // if(!info[0]->IsNumber() || !info[1]->IsNumber() || !info[2]->IsNumber()) {
  //   return Nan::ThrowError(Nan::New("Vector::New - expected arguments to be numbers").ToLocalChecked());
  // }

  std::string temp = *v8::String::Utf8Value(info[0]->ToString());
  std::wstring filename = s2ws(temp);

  if (!OpenClipboard(NULL))
  {
    return;
  }

  HANDLE hBitmap = GetClipboardData(CF_BITMAP);
  if (hBitmap != NULL)
  {
    // COM Fun
    HRESULT hr = CoInitialize(NULL);
    IWICImagingFactory *ipFactory = NULL;
    hr = CoCreateInstance(CLSID_WICImagingFactory,
                          NULL,
                          CLSCTX_INPROC_SERVER,
                          IID_IWICImagingFactory,
                          reinterpret_cast<void **>(&ipFactory));
    if (SUCCEEDED(hr))
    {
      std::cout << "Created IWICImagingFactory\n";

      IWICBitmap *ipBitmap = NULL;
      hr = ipFactory->CreateBitmapFromHBITMAP(reinterpret_cast<HBITMAP>(hBitmap),
                                              NULL,
                                              WICBitmapIgnoreAlpha,
                                              &ipBitmap);
      if (SUCCEEDED(hr))
      {
        // Retrieve the width and height of the clipboard image
        UINT width, height;
        hr = ipBitmap->GetSize(&width, &height);
        std::cout << "Converted into a IWICBitmap that is " << width << " by " << height << "\n";

        // Compute the output image width and height by constraining
        // the maximum width of the image to 800px.
        // TODO: parameterize, but hard code to 800px for now
        UINT final_width = 800;
        UINT output_width, output_height;
        float scaling_factor = (float)((float)final_width / (float)width);
        output_width = final_width;
        output_height = scaling_factor * height;

        std::cout << "resizing bitmap to: " << output_width << " by " << output_height << "\n";

        // Now resize it
        IWICBitmapScaler *ipScaler = NULL;
        hr = ipFactory->CreateBitmapScaler(&ipScaler);
        if (SUCCEEDED(hr))
        {
          std::cout << "Created an IWICBitmapScaler\n";

          hr = ipScaler->Initialize(ipBitmap,
                                    output_width,
                                    output_height,
                                    WICBitmapInterpolationModeHighQualityCubic);
          if (SUCCEEDED(hr))
          {
            std::cout << "Initialized it with the new output width and height\n";

            IWICBitmapEncoder *ipBitmapEncoder = NULL;
            hr = ipFactory->CreateEncoder(GUID_ContainerFormatPng,
                                          NULL,
                                          &ipBitmapEncoder);
            if (SUCCEEDED(hr))
            {
              std::cout << "Retrieved a PNG encoder\n";

              IWICStream *ipStream = NULL;
              hr = ipFactory->CreateStream(&ipStream);
              if (SUCCEEDED(hr))
              {
                std::cout << "Created an IWICStream from the encoder\n";
                hr = ipStream->InitializeFromFilename((filename + L".png").c_str(), GENERIC_WRITE);
                if (SUCCEEDED(hr))
                {
                  std::wcout << "Writing the image to this file: " << filename << ".png\n";
                  hr = ipBitmapEncoder->Initialize(ipStream, WICBitmapEncoderNoCache);
                  if (SUCCEEDED(hr))
                  {
                    std::cout << "Initialized the encoder with the stream and telling it to write to the image file!\n";
                    IWICBitmapFrameEncode *ipFrameEncoder = NULL;
                    hr = ipBitmapEncoder->CreateNewFrame(&ipFrameEncoder, NULL);
                    if (SUCCEEDED(hr))
                    {
                      std::cout << "Got a IWICBitmapFrameEncode interface\n";
                      hr = ipFrameEncoder->Initialize(NULL);
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
    }
    CoUninitialize();
  }
  else
  {
    std::cout << "No bitmap on clibpard\n";
  }

  return;
}