#include <nan.h>

class Clipboard : public Nan::ObjectWrap {
public:
    double x, y, z;

    static NAN_MODULE_INIT(Init);
    static NAN_METHOD(New);

    // Additional method to save with hard-coded filename
    static NAN_METHOD(Save);

    static Nan::Persistent<v8::FunctionTemplate> constructor;
};