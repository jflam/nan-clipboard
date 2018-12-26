#include <nan.h>

class Clipboard : public Nan::ObjectWrap {
public:
    static NAN_MODULE_INIT(Init);
    static NAN_METHOD(New);
    static NAN_METHOD(Save);

    static Nan::Persistent<v8::FunctionTemplate> constructor;
};