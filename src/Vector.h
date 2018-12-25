#include <nan.h>

class Vector : public Nan::ObjectWrap {
public:
    double x, y, z;

    static NAN_MODULE_INIT(Init);
    static NAN_METHOD(New);
    static NAN_METHOD(Add);

    static NAN_GETTER(HandleGetters);
    static NAN_SETTER(HandleSetters);

    // Additional method to save with hard-coded filename
    static NAN_METHOD(Save);

    static Nan::Persistent<v8::FunctionTemplate> constructor;
};