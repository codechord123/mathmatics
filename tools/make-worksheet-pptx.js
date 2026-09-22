const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';            // 10 x 5.625 in — Google Slides 기본과 동일
pres.author = '';
pres.title = '분수의 곱셈 - 나만의 문제 만들기';

const F = 'Noto Sans KR';
const NAVY='1F3864', INK='2F3A4A', MUTED='73809A', TEAL='0E7490';
const CARD='F5F7FA', LINE='DDE3EC', RULE='C6CEDC';

/* ---- {3/4} → ³⁄₄ , {1_1/2} → 1¹⁄₂ ---- */
function frac(out,n,d,b){
  out.push({text:n,options:Object.assign({},b,{superscript:true})});
  out.push({text:'⁄',options:Object.assign({},b)});
  out.push({text:d,options:Object.assign({},b,{subscript:true})});
}
function R(str,b){
  const out=[], re=/\{(\d+)_(\d+)\/(\d+)\}|\{(\d+)\/(\d+)\}/g;
  let last=0,m;
  while((m=re.exec(str))){
    if(m.index>last) out.push({text:str.slice(last,m.index),options:Object.assign({},b)});
    if(m[1]!==undefined){ out.push({text:m[1],options:Object.assign({},b)}); frac(out,m[2],m[3],b); }
    else frac(out,m[4],m[5],b);
    last=re.lastIndex;
  }
  if(last<str.length) out.push({text:str.slice(last),options:Object.assign({},b)});
  if(!out.length) out.push({text:str,options:Object.assign({},b)});
  return out;
}

/* ---- 치수 ---- */
const M=0.52, W=8.96;
const CY=1.04, CH=1.62;                  // 개념 정리 카드
const IX=M+0.22, IW=W-0.44;              // 카드 안쪽
const COLW=(IW-3*0.16)/4;
const COLX=[IX, IX+COLW+0.16, IX+2*(COLW+0.16), IX+3*(COLW+0.16)];
const BY=2.92, BH=2.26;                  // 답안 카드
const BW=4.34, BX=[M, M+BW+0.28];

function slide(o){
  const s = pres.addSlide();
  s.background = { color:'FFFFFF' };

  s.addText(R(o.title,{fontFace:F,fontSize:25,bold:true,color:NAVY}),
    {x:M,y:0.28,w:7.4,h:0.44,isTextBox:true,margin:0,valign:'middle'});
  s.addText(o.sub,
    {x:M,y:0.72,w:7.4,h:0.26,isTextBox:true,margin:0,valign:'middle',
     fontFace:F,fontSize:12,color:MUTED});
  s.addText(o.tag,
    {x:M+W-2.6,y:0.36,w:2.6,h:0.30,isTextBox:true,margin:0,align:'right',valign:'middle',
     fontFace:F,fontSize:11,bold:true,color:MUTED});

  /* 개념 정리 */
  s.addShape(pres.ShapeType.roundRect,
    {x:M,y:CY,w:W,h:CH,rectRadius:0.06,fill:{color:CARD},line:{color:LINE,width:1}});
  s.addText(o.boxLabel,
    {x:IX,y:CY+0.08,w:3,h:0.24,isTextBox:true,margin:0,valign:'middle',
     fontFace:F,fontSize:9.5,bold:true,color:NAVY});

  o.cols.forEach((c,i)=>{
    s.addText(R(c[0],{fontFace:F,fontSize:11.5,bold:true,color:NAVY}),
      {x:COLX[i],y:CY+0.36,w:COLW,h:0.24,isTextBox:true,margin:0,valign:'middle'});
    s.addText(R(c[1],{fontFace:F,fontSize:8.5,color:MUTED}),
      {x:COLX[i],y:CY+0.60,w:COLW,h:0.32,isTextBox:true,margin:0,valign:'top',lineSpacingMultiple:1.18});
    s.addText(R(c[2],{fontFace:F,fontSize:8.5,color:TEAL}),
      {x:COLX[i],y:CY+0.92,w:COLW,h:0.24,isTextBox:true,margin:0,valign:'middle'});
  });

  s.addShape(pres.ShapeType.line,
    {x:IX,y:CY+1.18,w:IW,h:0,line:{color:LINE,width:1}});

  const tipW=[2.90,2.98,IW-2.90-2.98-0.24];
  const tipX=[IX, IX+tipW[0]+0.12, IX+tipW[0]+tipW[1]+0.24];
  o.tips.forEach((t,i)=>{
    s.addText(
      [].concat(R(t[0],{fontFace:F,fontSize:9,bold:true,color:NAVY}),
                [{text:'   ',options:{fontFace:F,fontSize:9}}],
                R(t[1],{fontFace:F,fontSize:8.5,color:MUTED})),
      {x:tipX[i],y:CY+1.26,w:tipW[i],h:0.26,isTextBox:true,margin:0,valign:'middle'});
  });

  /* 답안 칸 */
  [['1. 내가 만든 문제',o.hint1],['2. 정답 및 풀이',o.hint2]].forEach((h,i)=>{
    s.addShape(pres.ShapeType.roundRect,
      {x:BX[i],y:BY,w:BW,h:BH,rectRadius:0.06,fill:{color:'FFFFFF'},line:{color:LINE,width:1}});
    s.addText(h[0],
      {x:BX[i]+0.22,y:BY+0.12,w:BW-0.44,h:0.26,isTextBox:true,margin:0,valign:'middle',
       fontFace:F,fontSize:12,bold:true,color:NAVY});
    s.addText(h[1],
      {x:BX[i]+0.22,y:BY+0.38,w:BW-0.44,h:0.22,isTextBox:true,margin:0,valign:'middle',
       fontFace:F,fontSize:8.5,color:MUTED});
    for(let k=0;k<4;k++){
      s.addShape(pres.ShapeType.line,
        {x:BX[i]+0.22,y:BY+0.86+k*0.40,w:BW-0.44,h:0,line:{color:RULE,width:0.75}});
    }
  });

  s.addNotes(o.notes);
  return s;
}

const TIPS_MAKE=[
  ['상황 넣기','이야기가 있는 문제로 만들어요'],
  ['수 고르기','약분이 되는 수를 넣으면 답이 깔끔해요'],
  ['답 먼저','답을 정해 놓고 거꾸로 만들어요']
];

slide({
  title:'분수의 곱셈', sub:'나만의 문제 만들기', tag:'5학년 1학기',
  boxLabel:'개념 정리',
  cols:[
    ['분수 × 자연수','분모는 그대로 두고\n분자에만 자연수를 곱해요','예) {3/4} × 2 = {6/4} = {1_1/2}'],
    ['자연수 × 분수','자연수와 분자를 곱하고\n분모로 나눠요','예) 6 × {2/3} = {12/3} = 4'],
    ['분수 × 분수','분자는 분자끼리,\n분모는 분모끼리 곱해요','예) {1/2} × {3/5} = {3/10}'],
    ['대분수의 곱셈','대분수를 가분수로 바꾼 뒤\n곱해요','예) {1_1/2} × 2 = {3/2} × 2 = 3']
  ],
  tips:[
    ['약분 먼저','곱하기 전에 약분하면 계산이 쉬워요'],
    ['가분수로','대분수는 반드시 가분수로 바꿔요'],
    ['답 정리','가분수는 대분수로, 기약분수로']
  ],
  hint1:'분수의 곱셈이 들어가는 이야기 문제를 만들어 보세요.',
  hint2:'식을 쓰고, 어떻게 풀었는지 말로 설명해 보세요.',
  notes:'분수의 곱셈 단원 전체를 정리한 문제 만들기 활동지입니다. 학생마다 사본을 배부해 사용하세요.'
});

slide({
  title:'(진)분수 × 자연수', sub:'나만의 문제 만들기', tag:'2 · 3차시',
  boxLabel:'개념 정리',
  cols:[
    ['계산 방법','분모는 그대로 두고\n분자에만 자연수를 곱해요','{3/4} × 2 = {6/4}'],
    ['약분하기','곱하기 전에 해도 되고\n곱한 뒤에 해도 돼요','{6/4} = {3/2}'],
    ['답 정리','가분수가 되면\n대분수로 바꿔 써요','{3/2} = {1_1/2}'],
    ['대분수일 때','가분수로 바꾼 뒤\n자연수를 곱해요','{1_1/3} × 3 = {4/3} × 3 = 4']
  ],
  tips:TIPS_MAKE,
  hint1:'"똑같은 양을 여러 번" 이야기로 만들어 보세요.',
  hint2:'식을 쓰고, 어떻게 풀었는지 말로 설명해 보세요.',
  notes:'2·3차시 (진)분수 × 자연수, 대분수 × 자연수용 활동지입니다.'
});

slide({
  title:'자연수 × (진)분수', sub:'나만의 문제 만들기', tag:'4 · 5차시',
  boxLabel:'개념 정리',
  cols:[
    ['계산 방법','자연수와 분자를 곱하고\n분모로 나눠요','6 × {2/3} = {12/3} = 4'],
    ['뜻 읽기','"6의 {2/3} 만큼"\n이라고 읽어요','6 ÷ 3 × 2 = 4'],
    ['크기 어림','진분수를 곱하면\n원래 수보다 작아져요','6 × {2/3} < 6'],
    ['대분수일 때','가분수로 바꾼 뒤\n곱해요','4 × {1_1/2} = 4 × {3/2} = 6']
  ],
  tips:[
    ['상황 넣기','"전체의 얼마만큼"으로 만들어요'],
    ['수 고르기','자연수가 분모의 배수면 답이 자연수'],
    ['답 먼저','답을 정해 놓고 거꾸로 만들어요']
  ],
  hint1:'"전체의 얼마만큼" 또는 "몇 배" 이야기로 만들어 보세요.',
  hint2:'식을 쓰고, 어떻게 풀었는지 말로 설명해 보세요.',
  notes:'4·5차시 자연수 × (진)분수, 자연수 × 대분수용 활동지입니다.'
});

slide({
  title:'분수 × 분수', sub:'나만의 문제 만들기', tag:'6 · 7차시',
  boxLabel:'개념 정리',
  cols:[
    ['계산 방법','분자는 분자끼리,\n분모는 분모끼리 곱해요','{1/2} × {3/5} = {3/10}'],
    ['약분 먼저','곱하기 전에 약분하면\n수가 작아져 쉬워요','{2/3} × {3/4} = {2/4} = {1/2}'],
    ['넓이로 보기','직사각형의\n가로 × 세로와 같아요','{1/2} m × {3/5} m = {3/10} m²'],
    ['크기 어림','진분수끼리 곱하면\n두 수보다 작아져요','{1/2} × {3/5} < {1/2}']
  ],
  tips:[
    ['상황 넣기','"넓이" 또는 "전체의 얼마의 얼마"'],
    ['수 고르기','약분이 되는 수를 넣으면 답이 깔끔해요'],
    ['답 먼저','답을 정해 놓고 거꾸로 만들어요']
  ],
  hint1:'"넓이" 또는 "전체의 얼마의 얼마" 이야기로 만들어 보세요.',
  hint2:'식을 쓰고, 어떻게 풀었는지 말로 설명해 보세요.',
  notes:'6·7차시 진분수 × 진분수, 넓이 맥락용 활동지입니다.'
});

slide({
  title:'대분수의 곱셈', sub:'나만의 문제 만들기', tag:'8차시',
  boxLabel:'개념 정리',
  cols:[
    ['계산 방법','두 수 모두 가분수로\n바꾼 뒤 곱해요','{1_1/2} × {2_1/3} = {3/2} × {7/3}'],
    ['약분하기','가분수로 바꾼 뒤\n약분해요','{21/6} = {7/2}'],
    ['답 정리','마지막에 대분수로\n바꿔 써요','{7/2} = {3_1/2}'],
    ['자주 하는 실수','자연수끼리, 분수끼리\n따로 곱하면 틀려요','{1_1/2} × {2_1/3} ≠ {2_1/6}']
  ],
  tips:TIPS_MAKE,
  hint1:'대분수가 두 번 들어가는 이야기로 만들어 보세요.',
  hint2:'가분수로 바꾸는 과정까지 함께 써 보세요.',
  notes:'8차시 대분수 × 대분수용 활동지입니다.'
});

slide({
  title:'어떤 상황으로 만들까?', sub:'분수의 곱셈 · 문제 만들기 도움판', tag:'도움판',
  boxLabel:'상황 고르기',
  cols:[
    ['똑같이 여러 번','같은 양을\n몇 번 더할 때','예) {3/4} L씩 3병이면?'],
    ['몇 배 구하기','어떤 양의\n몇 배를 구할 때','예) 6 kg의 {1_1/2} 배는?'],
    ['전체의 얼마만큼','전체에서\n일부를 구할 때','예) 12명의 {2/3} 는?'],
    ['직사각형의 넓이','가로 × 세로를\n구할 때','예) {1/2} m × {3/5} m']
  ],
  tips:[
    ['하나만 묻기','묻는 것이 한 가지여야 해요'],
    ['수가 알맞게','5학년이 풀 수 있는 수로'],
    ['답이 있게','꼭 답을 구할 수 있어야 해요']
  ],
  hint1:'위 네 가지 중 하나를 골라 이야기를 만들어 보세요.',
  hint2:'식을 쓰고, 어떻게 풀었는지 말로 설명해 보세요.',
  notes:'문제를 만들기 어려워하는 학생에게 먼저 보여 주는 도움판입니다.'
});

pres.writeFile({fileName:'분수의_곱셈_문제만들기.pptx'}).then(f=>console.log('wrote',f));
