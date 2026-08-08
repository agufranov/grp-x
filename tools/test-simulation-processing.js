'use strict';
const assert=require('assert');
const DSP=require('../simulation-processing.js');

const sampled=Array.from(DSP.resampleUniform(Float64Array.from([0,10,20]),1,.5));
assert.deepStrictEqual(sampled,[0,5,10,15,20],'uniform resampling must interpolate and preserve endpoints');

const raw=DSP.stabilizedLogSpectrum(Float64Array.from([1,0,0,0]),{floorRatio:1e-3,smoothBins:1});
assert.ok(raw.every(Number.isFinite),'relative spectral floor must keep logarithms finite');
assert.ok(Math.abs(raw[1]-Math.log(1e-3))<1e-12,'floor must be relative to the strongest bin');
const smooth=DSP.stabilizedLogSpectrum(Float64Array.from([1,0,0,0]),{floorRatio:1e-3,smoothBins:5});
assert.ok(smooth[0]<raw[0]&&smooth[1]>raw[1],'triangular smoothing must reduce a one-bin discontinuity');

const lifted=DSP.lifterCepstrum([{x:0,y:1},{x:1,y:1},{x:2,y:1},{x:3,y:1},{x:4,y:1}],{min:1,max:3,taper:1});
assert.deepStrictEqual(lifted.map(point=>point.x),[0,1,2,3,4],'lifter must retain its cosine taper support');
assert.strictEqual(lifted[0].y,0,'lower taper edge must reach zero');
assert.strictEqual(lifted[2].y,1,'pass band must remain unchanged');
assert.strictEqual(lifted.at(-1).y,0,'upper taper edge must reach zero');

const normalized=DSP.normalizeCepstrum([{x:0,y:-4},{x:1,y:2}]);
assert.deepStrictEqual(normalized,[{x:0,y:-1},{x:1,y:.5}],'cepstrum normalization must preserve shape and sign');

console.log('Simulation processing tests passed: resampling, spectral stabilization, lifter and normalization');
