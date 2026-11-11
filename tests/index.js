// Newer versions of Node (22 and up) read "node --test tests/" as "run the file
// tests/index.js", not "run everything in the tests folder". This file is that
// index.js: it simply runs the real tests.
require("./combine.test.js");
