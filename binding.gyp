{
  "targets": [{
    "target_name": "myModule",
    "include_dirs" : [
      "src",
      "<!(node -e \"require('nan')\")"
    ],
    "libraries": [
      "C:/Program Files (x86)/Windows Kits/10/Lib/10.0.17763.0/um/x64/windowscodecs.lib"
    ],
    "sources": [
      "src/index.cc",
      "src/Clipboard.cc"
    ]
  }]
}