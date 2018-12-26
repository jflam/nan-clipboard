// this is how we will require our module
const m = require('./')

const vec1 = new m.Clipboard(20, 10, 0)
vec1.save('hi')
console.log("made it here")