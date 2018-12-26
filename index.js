// this is how we will require our module
const m = require('./')

const vec1 = new m.Clipboard()
vec1.writeBitmapToDisk('hi')
console.log("made it here")