(function(root,factory){
  const api=factory();
  if(typeof module==='object'&&module.exports)module.exports=api;
  else root.GrpInterpolation=api;
})(typeof globalThis!=='undefined'?globalThis:this,function(){
  'use strict';
  const DEFAULT_METHOD='pchip',SUPPORTED_METHODS=Object.freeze(['pchip','linear']);

  function endpointSlope(here,next,deltaHere,deltaNext){
    let slope=((2*here+next)*deltaHere-here*deltaNext)/(here+next);
    if(Math.sign(slope)!==Math.sign(deltaHere))return 0;
    if(Math.sign(deltaHere)!==Math.sign(deltaNext)&&Math.abs(slope)>Math.abs(3*deltaHere))slope=3*deltaHere;
    return slope;
  }

  function pchipSlopes(points){
    const n=points.length,slopes=new Float64Array(n);
    if(n<2)return slopes;
    const h=new Float64Array(n-1),delta=new Float64Array(n-1);
    for(let i=0;i<n-1;i++){h[i]=points[i+1].x-points[i].x;delta[i]=(points[i+1].y-points[i].y)/h[i];}
    if(n===2){slopes[0]=delta[0];slopes[1]=delta[0];return slopes;}
    slopes[0]=endpointSlope(h[0],h[1],delta[0],delta[1]);
    for(let i=1;i<n-1;i++){
      const before=delta[i-1],after=delta[i];
      if(!before||!after||Math.sign(before)!==Math.sign(after)){slopes[i]=0;continue;}
      const beforeWeight=2*h[i]+h[i-1],afterWeight=h[i]+2*h[i-1];
      slopes[i]=(beforeWeight+afterWeight)/(beforeWeight/before+afterWeight/after);
    }
    slopes[n-1]=endpointSlope(h[n-2],h[n-3],delta[n-2],delta[n-3]);
    return slopes;
  }

  function createInterpolator(points,method=DEFAULT_METHOD){
    if(!SUPPORTED_METHODS.includes(method))throw Error(`Неизвестный метод интерполяции: ${method}`);
    const slopes=method==='pchip'?pchipSlopes(points):null,lastInterval=Math.max(0,points.length-2);
    return function interpolate(x,interval){
      const index=Math.max(0,Math.min(lastInterval,interval)),a=points[index],b=points[Math.min(index+1,points.length-1)],span=b.x-a.x;
      if(!span)return a.y;
      const t=Math.max(0,Math.min(1,(x-a.x)/span));
      if(method==='linear')return a.y+t*(b.y-a.y);
      const t2=t*t,t3=t2*t;
      return(2*t3-3*t2+1)*a.y+(t3-2*t2+t)*span*slopes[index]+(-2*t3+3*t2)*b.y+(t3-t2)*span*slopes[index+1];
    };
  }

  return Object.freeze({DEFAULT_METHOD,SUPPORTED_METHODS,createInterpolator});
});
