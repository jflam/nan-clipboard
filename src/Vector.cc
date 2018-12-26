#include "Vector.h"
#include <comdef.h>
#include <windows.h>
#include <wincodec.h>

using namespace v8;

Nan::Persistent<v8::FunctionTemplate> Vector::constructor;

NAN_MODULE_INIT(Vector::Init) {
  v8::Local<v8::FunctionTemplate> ctor = Nan::New<v8::FunctionTemplate>(Vector::New);
  constructor.Reset(ctor);
  ctor->InstanceTemplate()->SetInternalFieldCount(1);
  ctor->SetClassName(Nan::New("Vector").ToLocalChecked());

  // link our getters and setter to the object property
  Nan::SetAccessor(ctor->InstanceTemplate(), Nan::New("x").ToLocalChecked(), Vector::HandleGetters, Vector::HandleSetters);
  Nan::SetAccessor(ctor->InstanceTemplate(), Nan::New("y").ToLocalChecked(), Vector::HandleGetters, Vector::HandleSetters);
  Nan::SetAccessor(ctor->InstanceTemplate(), Nan::New("z").ToLocalChecked(), Vector::HandleGetters, Vector::HandleSetters);

  Nan::SetPrototypeMethod(ctor, "add", Add);
  Nan::SetPrototypeMethod(ctor, "save", Save);

  target->Set(Nan::New("Vector").ToLocalChecked(), ctor->GetFunction());
}

NAN_METHOD(Vector::New) {

  // throw an error if constructor is called without new keyword
  if(!info.IsConstructCall()) {
    return Nan::ThrowError(Nan::New("Vector::New - called without new keyword").ToLocalChecked());
  }

  // expect exactly 3 arguments
  if(info.Length() != 3) {
    return Nan::ThrowError(Nan::New("Vector::New - expected arguments x, y, z").ToLocalChecked());
  }

  // expect arguments to be numbers
  if(!info[0]->IsNumber() || !info[1]->IsNumber() || !info[2]->IsNumber()) {
    return Nan::ThrowError(Nan::New("Vector::New - expected arguments to be numbers").ToLocalChecked());
  }

  // create a new instance and wrap our javascript instance
  Vector* vec = new Vector();
  vec->Wrap(info.Holder());

  // initialize it's values
  vec->x = info[0]->NumberValue();
  vec->y = info[1]->NumberValue();
  vec->z = info[2]->NumberValue();

  // return the wrapped javascript instance
  info.GetReturnValue().Set(info.Holder());
}

NAN_METHOD(Vector::Add) {
  // unwrap this Vector
  Vector * self = Nan::ObjectWrap::Unwrap<Vector>(info.This());

  if (!Nan::New(Vector::constructor)->HasInstance(info[0])) {
    return Nan::ThrowError(Nan::New("Vector::Add - expected argument to be instance of Vector").ToLocalChecked());
  }
  // unwrap the Vector passed as argument
  Vector * otherVec = Nan::ObjectWrap::Unwrap<Vector>(info[0]->ToObject());

  // specify argument counts and constructor arguments
  const int argc = 3;
  v8::Local<v8::Value> argv[argc] = {
    Nan::New(self->x + otherVec->x),
    Nan::New(self->y + otherVec->y),
    Nan::New(self->z + otherVec->z)
  };

  // get a local handle to our constructor function
  v8::Local<v8::Function> constructorFunc = Nan::New(Vector::constructor)->GetFunction();
  // create a new JS instance from arguments
  v8::Local<v8::Object> jsSumVec = Nan::NewInstance(constructorFunc, argc, argv).ToLocalChecked();

  info.GetReturnValue().Set(jsSumVec);
}


NAN_GETTER(Vector::HandleGetters) {
  Vector* self = Nan::ObjectWrap::Unwrap<Vector>(info.This());

  std::string propertyName = std::string(*Nan::Utf8String(property));
  if (propertyName == "x") {
    info.GetReturnValue().Set(self->x);
  } else if (propertyName == "y") {
    info.GetReturnValue().Set(self->y);
  } else if (propertyName == "z") {
    info.GetReturnValue().Set(self->z);
  } else {
    info.GetReturnValue().Set(Nan::Undefined());
  }
}

NAN_SETTER(Vector::HandleSetters) {
  Vector* self = Nan::ObjectWrap::Unwrap<Vector>(info.This());

  if(!value->IsNumber()) {
    return Nan::ThrowError(Nan::New("expected value to be a number").ToLocalChecked());
  }

  std::string propertyName = std::string(*Nan::Utf8String(property));
  if (propertyName == "x") {
    self->x = value->NumberValue();
  } else if (propertyName == "y") {
    self->y = value->NumberValue();
  } else if (propertyName == "z") {
    self->z = value->NumberValue();
  }
}

NAN_METHOD(Vector::Save) {

    if (!OpenClipboard(NULL)) {
        return;
    }

    HANDLE hBitmap = GetClipboardData(CF_BITMAP);
    if (hBitmap != NULL) {
      // Get dimensions of the bitmap
      BITMAP bm;
      GetObject(hBitmap, sizeof(bm), &bm);
      char buffer[255];
      sprintf(buffer, "width = %i, height = %i\n", bm.bmWidth, bm.bmHeight);
      printf(buffer);

      // COM Fun
      HRESULT hr = CoInitialize(NULL);
      IWICImagingFactory* ipFactory = NULL;
      hr = CoCreateInstance(CLSID_WICImagingFactory, 
                            NULL, 
                            CLSCTX_INPROC_SERVER, 
                            IID_IWICImagingFactory, 
                            reinterpret_cast<void**>(&ipFactory));
      if (SUCCEEDED(hr)) {
        printf("Got a reference to the WIC component\n");

        IWICBitmap* ipBitmap = NULL;
        hr = ipFactory->CreateBitmapFromHBITMAP(reinterpret_cast<HBITMAP>(hBitmap), 
                                                NULL, 
                                                WICBitmapIgnoreAlpha, 
                                                &ipBitmap);
        if (SUCCEEDED(hr)) {
          printf("Converted into a IWICBitmap!\n");

          // Retrieve the width and height of the clipboard image
          UINT width, height;
          hr = ipBitmap->GetSize(&width, &height);
          sprintf(buffer, "width = %i, height = %i\n", width, height);
          printf(buffer);

          // Compute the output image width and height by constraining
          // the maximum width of the image to 800px.
          // TODO: parameterize, but hard code to 800px for now
          UINT final_width = 800;
          UINT output_width, output_height;
          float scaling_factor = (float)((float)final_width / (float)width);
          output_width = final_width;
          output_height = scaling_factor * height;

          // Now let's get a PngEncoder
          IWICBitmapEncoder* ipBitmapEncoder = NULL;
          hr = ipFactory->CreateEncoder(GUID_ContainerFormatPng,
                                        NULL,
                                        &ipBitmapEncoder);
          if (SUCCEEDED(hr)) {
            printf("Got a reference to a PNG encoder\n");

            IWICStream* ipStream = NULL;
            hr = ipFactory->CreateStream(&ipStream);
            if (SUCCEEDED(hr)) {
              printf("Created an IWICStream from the encoder\n");
              hr = ipStream->InitializeFromFilename(L"./output.png", GENERIC_WRITE);
              if (SUCCEEDED(hr)) {
                printf("Initialized with a backing file\n");
                hr = ipBitmapEncoder->Initialize(ipStream, WICBitmapEncoderNoCache);
                if (SUCCEEDED(hr)) {
                  printf("Initialized the encoder with the stream and telling it to write to the image file!\n");
                  IWICBitmapFrameEncode* ipFrameEncoder = NULL;
                  hr = ipBitmapEncoder->CreateNewFrame(&ipFrameEncoder, NULL);
                  if (SUCCEEDED(hr)) {
                    printf("Got a IWICBitmapFrameEncode interface\n");
                    hr = ipFrameEncoder->Initialize(NULL);
                    if (SUCCEEDED(hr)) {
                      printf("Initialized the frame encoder\n");
                      hr = ipFrameEncoder->SetSize(bm.bmWidth, bm.bmHeight);
                      if (SUCCEEDED(hr))
                      {
                        printf("set size succeeded");
                        // TODO: const_cast?
                        WICPixelFormatGUID formatGuid = GUID_WICPixelFormat24bppBGR;
                        hr = ipFrameEncoder->SetPixelFormat(&formatGuid);
                        if (SUCCEEDED(hr))
                        {
                          printf("set pixel format to RGB 8 bits per channel\n");
                          hr = ipFrameEncoder->WriteSource(ipBitmap, NULL);
                          if (SUCCEEDED(hr))
                          {
                            printf("set write source to the IWICBitamp object\n");
                            hr = ipFrameEncoder->Commit();
                            if (SUCCEEDED(hr))
                            {
                              printf("committed the frame encoder\n");
                              hr = ipBitmapEncoder->Commit();
                              if (SUCCEEDED(hr))
                              {
                                printf("committed the bitmap encoder\n");
                              }
                            }
                          }
                        }
                      }
                      else
                      {
                        printf("uh oh\n");
                        _com_error err(hr);
                        LPCTSTR errMsg = err.ErrorMessage();
                        printf(errMsg);
                      }
                    }
                  }
                }
              }
            }
          }
        } else {
          _com_error err(hr);
          LPCTSTR errMsg = err.ErrorMessage();
          printf(errMsg);
        }
        ipFactory->Release();
      }
      CoUninitialize();
    } else {
      printf("No bitmap on clibpard\n");
    }

    return;
}