(function(root,factory){
  const api=factory();
  if(typeof module==='object'&&module.exports)module.exports=api;
  if(root)root.GrpSimulationProcessing=api;
})(typeof globalThis!=='undefined'?globalThis:this,function(){
  'use strict';

  function resampleUniform(values,sourceStep,targetStep){
    if(!values?.length||!(sourceStep>0)||!(targetStep>0))return new Float64Array();
    const duration=(values.length-1)*sourceStep,count=Math.floor(duration/targetStep)+1,result=new Float64Array(count);
    for(let i=0;i<count;i++){
      const position=i*targetStep/sourceStep,left=Math.floor(position),right=Math.min(values.length-1,left+1),fraction=position-left;
      result[i]=values[left]+(values[right]-values[left])*fraction;
    }
    return result;
  }

  function stabilizedLogSpectrum(magnitudes,{floorRatio=0,smoothBins=1}={}){
    const n=magnitudes?.length||0,result=new Float64Array(n);if(!n)return result;
    let maximum=0;for(const magnitude of magnitudes)if(Number.isFinite(magnitude)&&magnitude>maximum)maximum=magnitude;
    const floor=Math.max(1e-12,maximum*Math.max(0,floorRatio));
    for(let i=0;i<n;i++)result[i]=Math.log(Math.max(floor,Number.isFinite(magnitudes[i])?magnitudes[i]:0));
    const size=Math.max(1,Math.floor(smoothBins));if(size<3)return result;
    const radius=Math.floor(Math.min(size,n-1)/2);if(!radius)return result;
    const smoothed=new Float64Array(n),normalizer=(radius+1)*(radius+1);
    for(let i=0;i<n;i++){
      let sum=0;for(let offset=-radius;offset<=radius;offset++){const weight=radius+1-Math.abs(offset),index=(i+offset+n)%n;sum+=result[index]*weight;}
      smoothed[i]=sum/normalizer;
    }
    return smoothed;
  }

  function lifterCepstrum(points,{min=0,max=Infinity,taper=0,domainMax=max+taper}={}){
    if(!Array.isArray(points)||!points.length||!(max>min))return[];
    const edge=Math.max(0,taper),from=Math.max(0,min-edge),to=max+edge,result=[];
    for(const point of points){
      if(!(point.x>=0&&point.x<=domainMax))continue;
      let weight=point.x>=from&&point.x<=to?1:0;
      if(edge>0&&point.x>=from&&point.x<min)weight=.5-.5*Math.cos(Math.PI*(point.x-from)/edge);
      else if(edge>0&&point.x>max&&point.x<=to)weight=.5+.5*Math.cos(Math.PI*(point.x-max)/edge);
      result.push({x:point.x,y:point.y*weight});
    }
    return result;
  }

  function normalizeCepstrum(points){
    if(!Array.isArray(points)||!points.length)return[];
    let scale=0;for(const point of points)if(Number.isFinite(point.y))scale=Math.max(scale,Math.abs(point.y));
    if(!(scale>0))return points.map(point=>({x:point.x,y:0}));
    return points.map(point=>({x:point.x,y:point.y/scale}));
  }

  return{resampleUniform,stabilizedLogSpectrum,lifterCepstrum,normalizeCepstrum};
});
