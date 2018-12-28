#include <nan.h>

class Clipboard : public Nan::ObjectWrap {
public:
    static NAN_MODULE_INIT(Init);
    static NAN_METHOD(New);
    static NAN_METHOD(WriteBitmapToDisk);

    static Nan::Persistent<v8::FunctionTemplate> constructor;
};