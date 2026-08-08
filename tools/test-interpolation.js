'use strict';
const assert=require('node:assert/strict');
const {createInterpolator}=require('../interpolation.js');

const close=(actual,expected,tolerance=1e-10)=>assert.ok(Math.abs(actual-expected)<=tolerance,`${actual} != ${expected}`);
const evaluate=(points,method,x)=>{
  let interval=0;
  while(interval+1<points.length&&points[interval+1].x<x)interval++;
  return createInterpolator(points,method)(x,interval);
};

const line=[{x:0,y:1},{x:.3,y:1.6},{x:1.4,y:3.8},{x:4,y:9}];
for(const method of ['linear','pchip'])for(const x of [0,.1,.3,.8,1.4,2.7,4])close(evaluate(line,method,x),1+2*x);

const shaped=[{x:0,y:0},{x:.2,y:2},{x:1.7,y:1},{x:4,y:3}];
const pchip=createInterpolator(shaped,'pchip');
for(let interval=0;interval<shaped.length-1;interval++){
  const a=shaped[interval],b=shaped[interval+1],low=Math.min(a.y,b.y),high=Math.max(a.y,b.y);
  for(let i=0;i<=100;i++){
    const value=pchip(a.x+(b.x-a.x)*i/100,interval);
    assert.ok(value>=low-1e-12&&value<=high+1e-12,`PCHIP overshoot on interval ${interval}: ${value}`);
  }
  close(pchip(a.x,interval),a.y);close(pchip(b.x,interval),b.y);
}

assert.throws(()=>createInterpolator(line,'cubic'),/Неизвестный метод/);
console.log('Interpolation tests passed: linear reference and shape-preserving PCHIP on irregular grids');
