/*
tests for spServerJsom.js
6/6/17

each test will run and log the function name and true if passes
or function name and false if fails
*/

import * as spu from 'pd-sputil';
import * as spa from './src/library.js';

var $ = require('jquery');

$.noConflict();

//ajax testing list id - ef3baa99-88f1-4116-a524-66d2dc6f08bf

const testProcess = (function() {
    var objProto = {
        init: function() {

            let self = this;
        }
    }; 

    return function() {
        var obj = Object.create(objProto);
        return obj;
    };
})();

spu.domReady(function() {
    testProcess().init();
});
