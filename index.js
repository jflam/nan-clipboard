// this is how we will require our module
const m = require('./')

const vec1 = new m.Clipboard()
vec1.writeBitmapToDisk('filename', 'png', 800, true)
console.log("made it here")