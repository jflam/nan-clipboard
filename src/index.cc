#include <nan.h>
#include "Clipboard.h"

NAN_MODULE_INIT(InitModule) {
    Clipboard::Init(target);
}

NODE_MODULE(myModule, InitModule);