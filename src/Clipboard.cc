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

// TODO: how create a destructor to clean up resources in the Clipboard object?
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

// Helper method to convert a UTF-8 encoded string into a STL wstring
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

// Helper method that generates an error message + 32 bit hex representation of the HRESULT
void RaiseError(std::string errorMessage, HRESULT hr)
{
  std::stringstream ss;
  ss << errorMessage << "0x" << std::uppercase << std::setfill('0') << std::setw(8) << std::hex << hr;
  Nan::ThrowError(Nan::New(ss.str()).ToLocalChecked());
}

// Internal method that is used to do the work on the validated / parsed / decoded parameters 
void InternalWriteBitmapToDisk(std::wstring filename, std::wstring file_format, int width_constraint, bool write_full)
{
  HBITMAP hBitmap;
  HRESULT hr;

  IWICImagingFactory *ipFactory = NULL;
  IWICBitmapScaler *ipScaler = NULL;
  IWICBitmapEncoder *ipBitmapEncoder = NULL;
  IWICStream *ipStream = NULL;
  IWICBitmapFrameEncode *ipFrameEncoder = NULL;
  IWICBitmap *ipBitmap = NULL;

  // Attempt to open the Clipboard object. Failure will result in exit. A successful call must be balanced by
  // a call to CloseClipboard()
  if (!OpenClipboard(NULL))
  {
    return;
  }

  // Attempt to retrieve a HBITMAP object from the clipboard. Failure is OK - there is no bitmap to write to disk.
  hBitmap = reinterpret_cast<HBITMAP>(GetClipboardData(CF_BITMAP));
  if (hBitmap == NULL)
  {
    goto FreeClipboard;
  }

  // Must initialize COM on this thread before we can use it. A successful call must be balanced by a call to
  // CoUninitialize() which will unload inproc servers from the process as part of cleanup. TODO: We may want to
  // have an explicit init / uninit protocol (perhaps with the creation of the Clipboard object) in the future
  // so we don't thrash the OS loading and unloading the WIC module.
  hr = CoInitialize(NULL);
  if (FAILED(hr))
  {
    RaiseError("Failed to initialize COM: ", hr);
    goto FreeClipboard;
  }

  // The WIC Imaging Factory object is used to create resources that we will be using:
  // 1. A WIC Bitmap object that we create from the HBITMAP object that we received from the clipboard
  // 2. A WIC Bitmap Scaler object that we will use to resize the bitmap to the size determined by width_constraint
  // 3. A WIC Stream object that is used to write the serialized bitmap(s) to disk
  // 4. A WIC Bitmap Encoder (either for PNG or JPG) that is used to serialize the bitmap(s) to disk
  hr = CoCreateInstance(CLSID_WICImagingFactory,
                        NULL,
                        CLSCTX_INPROC_SERVER,
                        IID_IWICImagingFactory,
                        reinterpret_cast<void **>(&ipFactory));
  if (FAILED(hr))
  {
    RaiseError("Failed to initialize WIC Imaging Factory object: ", hr);

    // Notice that COM is initialized here, so we go to the section where we release all COM interfaces
    // that have been explicitly assigned
    goto FreeCOM;
  }

  // Bitmap objects stored on the clipboard are referenced by a generic Windows HANDLE object. Here we 
  // initialize the WIC Bitmap object using the data stored in the HBITMAP.
  hr = ipFactory->CreateBitmapFromHBITMAP(reinterpret_cast<HBITMAP>(hBitmap),
                                          NULL,
                                          WICBitmapIgnoreAlpha,
                                          &ipBitmap);
  if (FAILED(hr))
  {
    RaiseError("Failed to construct a WIC Bitmap object from the HBITMAP from the clipboard: ", hr);
    goto FreeCOM;
  }

  // Retrieve the width and height of the clipboard image
  UINT width, height;
  hr = ipBitmap->GetSize(&width, &height);
  if (FAILED(hr))
  {
    RaiseError("Could not get the width and height of the WIC Bitmap object: ", hr);
    goto FreeCOM;
  }

  // Compute the output image width and height by constraining
  // the maximum width of the image to 800px.
  UINT output_width, output_height;
  float scaling_factor = (float)((float)width_constraint / (float)width);
  output_width = width_constraint;
  output_height = scaling_factor * height;

  // Now resize it using a WIC Bitmap Scaler object
  hr = ipFactory->CreateBitmapScaler(&ipScaler);
  if (FAILED(hr)) 
  {
    RaiseError("Could not create a WIC Bitmap scaler object: ", hr);
    goto FreeCOM;
  }

  // Create a WIC Bitmap Scaler object that we will use to resize the object. In this case note that the
  // HighQualityCubic interpolation option only ships in Windows 10. TODO: perhaps have a fallback option
  // here in case the user is running on an older OS (we really need to test on an older OS to see if this is true)
  hr = ipScaler->Initialize(ipBitmap,
                            output_width,
                            output_height,
                            WICBitmapInterpolationModeHighQualityCubic);
  if (FAILED(hr))
  {
    RaiseError("Could not initialize the WIC Bitmap scaler object InterpolationMode High Quality Cubic: ", hr);
    goto FreeCOM;
  }

  // Create the appropriate WIC Bitmap Encoder object based on whether the user wants the file to be 
  // serialized as a PNG or a JPEG. TODO: in the future perhaps have an "auto" mode that attempts to serialize
  // to an in-memory buffer to see which mechanism results in smaller images? 
  // TODO: perhaps write a macro that randomly injects failing HRESULTs into debug builds to test the
  // recovery code paths?
  GUID encoderId = file_format == L"png" ? GUID_ContainerFormatPng : GUID_ContainerFormatJpeg;
  hr = ipFactory->CreateEncoder(encoderId,
                                NULL,
                                &ipBitmapEncoder);
  if (FAILED(hr))
  {
    RaiseError("Could not create the PNG or JPG encoder: ", hr);
    goto FreeCOM;
  }

  // Construct a WIC stream object using the factory
  hr = ipFactory->CreateStream(&ipStream);
  if (FAILED(hr)) {
    RaiseError("Could not create an IStream object: ", hr);
    goto FreeCOM;
  }

  // Initialize the WIC stream to write its contents out to the filename specified below. Ensure that
  // the correct filename extension (png|jpg) is appended
  hr = ipStream->InitializeFromFilename((filename + L"." + file_format).c_str(), GENERIC_WRITE);
  if (FAILED(hr))
  {
    RaiseError("Failed to initialize a writeable stream: ", hr);
    goto FreeCOM;
  }

  // Tell the WIC Bitmap Encoder to write to the WIC stream object
  hr = ipBitmapEncoder->Initialize(ipStream, WICBitmapEncoderNoCache);
  if (FAILED(hr))
  {
    RaiseError("Failed to initialize the bitmap encoder using the stream: ", hr);
    goto FreeCOM;
  }

  // Construct a WIC Frame Encoder object that we will be using to encode the WIC Bitmap object
  hr = ipBitmapEncoder->CreateNewFrame(&ipFrameEncoder, NULL);
  if (FAILED(hr))
  {
    RaiseError("Failed to create a new frame encoder using the bitmap encoder: ", hr);
    goto FreeCOM;
  }

  // Initialize the WIC Frame Encoder object. This is required otherwise subsequent calls will fail
  hr = ipFrameEncoder->Initialize(NULL);
  if (FAILED(hr))
  {
    RaiseError("Failed to initialize the frame encoder: ", hr);
    goto FreeCOM;
  }

  // Set the size of the frame encoder to be the final size of the image
  // TODO: when we do this twice - for original and resized image we may want to factor this code out 
  hr = ipFrameEncoder->SetSize(output_width, output_height);
  if (FAILED(hr))
  {
    RaiseError("Failed to set the output size for the frame encoder: ", hr);
    goto FreeCOM;
  }

  // Set the correct pixel format. This is RGB 8 bits per color channel.
  WICPixelFormatGUID formatGuid = GUID_WICPixelFormat24bppRGB;
  hr = ipFrameEncoder->SetPixelFormat(&formatGuid);
  if (FAILED(hr))
  {
    RaiseError("Failed to set the pixel format (WICPixelFormat24bppRGB) for the frame encoder: ", hr);
    goto FreeCOM;
  }

  // Tell the WIC Frame Encoder to use the WIC Bitmap Scaler object we created earlier.
  // This completes the construction of the pipeline:
  // Bitmap -> BitmapScaler -> PngEncoder -> FrameEncoder -> Stream
  hr = ipFrameEncoder->WriteSource(ipScaler, NULL);
  if (FAILED(hr))
  {
    RaiseError("Failed to set the write source of the frame encoder to the bitmap scaler: ", hr);
    goto FreeCOM;
  }

  // Tell the WIC Frame Encoder to serialize the frame to the stream
  hr = ipFrameEncoder->Commit();
  if (FAILED(hr)) 
  {
    RaiseError("Failed to commit the frame encoder: ", hr);
    goto FreeCOM;
  }

  // Tell the WIC Bitmap Encoder to serialize the image (which includes a frame) to the stream
  hr = ipBitmapEncoder->Commit();
  if (FAILED(hr))
  {
    RaiseError("Failed to commit the bitmap encoder: ", hr);
    goto FreeCOM;
  }

FreeCOM:
  // Free all COM interfaces that have been assigned. There should only be a single AddRef to any
  // of these interfaces, so the single Release if unassigned will do the right thing.
  if (ipStream != NULL)
  {
    ipStream->Release();
  }

  if (ipScaler != NULL)
  {
    ipScaler->Release();
  }

  if (ipBitmapEncoder != NULL)
  {
    ipBitmapEncoder->Release();
  }

  if (ipFrameEncoder != NULL)
  {
    ipFrameEncoder->Release();
  }

  if (ipFactory != NULL)
  {
    ipFactory->Release();
  }

  // Turn off COM for this thread - this will likely unload the WIC component, so TODO we really 
  // should move this to the destructor of the Clipboard class when we write that.
  CoUninitialize();

FreeClipboard:
  // Ensure that we have closed the clipboard from this thread to balance the call to OpenClipboard()
  CloseClipboard();
  return;
}

// TODO: should we put all error messages here?
const std::string writeBitmapToDiskParameterError{"writeBitmapToDisk - expected arguments filename, file_format, width_constraint, write_full"};

// Entry poit for the writeBitmapToDisk method that calls validates incoming parameters.
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
  std::wstring file_format = s2ws(*v8::String::Utf8Value(Isolate::GetCurrent(), info[1]->ToString()));
  // TODO: figure out how to deal with deprecated Int32Value method
  int width_constraint = info[2]->Int32Value();
  bool write_full = info[3]->BooleanValue();

  if (!(file_format == L"png" || file_format == L"jpg"))
  {
    return Nan::ThrowError(Nan::New("writeBitmapToDisk - file_format must be png|jpg").ToLocalChecked());
  }

  InternalWriteBitmapToDisk(filename, file_format, width_constraint, write_full);
}