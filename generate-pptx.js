const pptxgen = require('pptxgenjs');
const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'mihamix339-lab';
pptx.subject = 'Молодёжный эко-патруль порта';
pptx.title = 'Молодёжный эко-патруль порта';
pptx.company = 'Калининград';
pptx.lang = 'ru-RU';
pptx.theme = { headFontFace: 'Aptos Display', bodyFontFace: 'Aptos', lang: 'ru-RU' };
pptx.defineSlideMaster({
  title: 'MASTER',
  background: { color: 'F7FBFD' },
  objects: [
    { rect: { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: 'F7FBFD' }, line: { color: 'F7FBFD' } } },
    { line: { x: 0, y: 0.01, w: 13.333, h: 0, line: { color: 'D9E8EF', transparency: 55, width: 1 } } },
    { text: { text: 'ЭКОНОМИКА  ·  РЫНОК ТРУДА  ·  ESG', options: { x: .77, y: .38, w: 5, h: .22, fontFace: 'Aptos', fontSize: 8.5, bold: true, charSpacing: 1.2, color: '1689D5', margin: 0 } } },
    { text: { text: '01 / ФОРУМ', options: { x: 11.45, y: .38, w: 1.1, h: .2, fontSize: 8.5, bold: true, color: '8CA5B3', align: 'right', margin: 0 } } }
  ],
  slideNumber: { x: 12.8, y: 7.15, color: '9AB0BA', fontSize: 7 }
});
const slide = pptx.addSlide('MASTER');
const navy='12304A', blue='1689D5', muted='5F7487', line='D9E8EF', teal='0E7695';
slide.addShape(pptx.ShapeType.ellipse,{x:10.55,y:-1.2,w:4.6,h:4.6,fill:{color:'D9F4F1',transparency:12},line:{color:'D9F4F1',transparency:100}});
slide.addShape(pptx.ShapeType.ellipse,{x:-1.6,y:6.1,w:3.1,h:3.1,fill:{color:'E5F0FF',transparency:15},line:{color:'E5F0FF',transparency:100}});
slide.addText('Молодёжный ',{x:.77,y:.86,w:6.2,h:.58,fontFace:'Aptos Display',fontSize:32,bold:true,color:navy,margin:0,breakLine:false,fit:'shrink'});
slide.addText('эко‑патруль порта',{x:.77,y:1.38,w:6.5,h:.62,fontFace:'Aptos Display',fontSize:32,bold:true,color:blue,margin:0,fit:'shrink'});
slide.addText('От занятости — к устойчивому развитию. Калининградская модель по мотивам Sea Ranger Service, Роттердам.',{x:.8,y:2.13,w:7.15,h:.5,fontSize:13,color:muted,breakLine:false,margin:0,fit:'shrink'});
// left card
slide.addShape(pptx.ShapeType.roundRect,{x:.78,y:2.92,w:7.25,h:2.9,rectRadius:.12,fill:{color:'FFFFFF',transparency:6},line:{color:line,width:1},shadow:{type:'outer',color:'B7CCD6',opacity:.12,blur:2,angle:45,distance:2}});
slide.addText('Учимся через реальную работу',{x:1.08,y:3.2,w:5.7,h:.32,fontSize:19,bold:true,color:navy,margin:0});
slide.addText('Молодёжь 18–29 лет получает зарплату, морскую подготовку и первый опыт прямо в портовой среде.',{x:1.08,y:3.7,w:6.1,h:.62,fontSize:13.5,color:'36556A',breakLine:false,margin:0,fit:'shrink'});
slide.addShape(pptx.ShapeType.roundRect,{x:1.08,y:4.58,w:2.65,h:.42,rectRadius:.08,fill:{color:'E8F8F5'},line:{color:'E8F8F5'}});
slide.addShape(pptx.ShapeType.ellipse,{x:1.25,y:4.71,w:.1,h:.1,fill:{color:'20BFA9'},line:{color:'20BFA9'}});
slide.addText('Порт как площадка развития',{x:1.48,y:4.68,w:2.05,h:.14,fontSize:9.5,bold:true,color:'16786F',margin:0});
slide.addShape(pptx.ShapeType.roundRect,{x:1.08,y:5.2,w:2.72,h:.47,rectRadius:.06,fill:{color:'F1F8FB'},line:{color:'E2EEF2'}});
slide.addText('до 200',{x:1.25,y:5.27,w:1.2,h:.2,fontSize:21,bold:true,color:blue,margin:0}); slide.addText('трудоустроенных\nза 3 года',{x:2.42,y:5.25,w:1.15,h:.25,fontSize:8.2,color:muted,margin:0,fit:'shrink'});
slide.addShape(pptx.ShapeType.roundRect,{x:4.0,y:5.2,w:2.72,h:.47,rectRadius:.06,fill:{color:'F1F8FB'},line:{color:'E2EEF2'}});
slide.addText('50%',{x:4.17,y:5.27,w:.8,h:.2,fontSize:21,bold:true,color:blue,margin:0}); slide.addText('остаются в\nморском секторе',{x:4.98,y:5.25,w:1.35,h:.25,fontSize:8.2,color:muted,margin:0,fit:'shrink'});
// right visual card
slide.addShape(pptx.ShapeType.roundRect,{x:8.35,y:2.92,w:4.2,h:2.9,rectRadius:.12,fill:{color:'0E3E60'},line:{color:'0E3E60'}});
slide.addText('Три эффекта одной модели',{x:8.72,y:3.22,w:3.25,h:.28,fontSize:16,bold:true,color:'FFFFFF',margin:0}); slide.addText('работа  ·  кадры  ·  экология',{x:8.72,y:3.6,w:3,h:.2,fontSize:9.5,color:'B9E6ED',margin:0});
// stylized ship
slide.addShape(pptx.ShapeType.chevron,{x:9.1,y:4.54,w:2.65,h:.38,rotate:0,fill:{color:'FFFFFF'},line:{color:'FFFFFF'}});
slide.addShape(pptx.ShapeType.rect,{x:9.72,y:4.12,w:1.2,h:.42,fill:{color:'DFF8F2'},line:{color:'DFF8F2'}});
slide.addShape(pptx.ShapeType.line,{x:10.26,y:3.72,w:0,h:.74,line:{color:'DFF8F2',width:1.3}});
slide.addShape(pptx.ShapeType.triangle,{x:10.28,y:3.72,w:.5,h:.2,rotate:90,fill:{color:'FFB45C'},line:{color:'FFB45C'}});
[['Кадры',11.35,4.02],['Доход',8.7,4.45],['ESG',11.45,5.1]].forEach(([t,x,y])=>{slide.addShape(pptx.ShapeType.roundRect,{x,y,w:.78,h:.3,rectRadius:.06,fill:{color:'FFFFFF',transparency:85},line:{color:'FFFFFF',transparency:55,width:.6}});slide.addText(t,{x:x+.08,y:y+.09,w:.62,h:.1,fontSize:7.5,bold:true,color:'FFFFFF',align:'center',margin:0});});
// footer
slide.addText('ПАРТНЁРСТВО ДЛЯ ЗАПУСКА',{x:.8,y:6.32,w:2.3,h:.16,fontSize:8.5,bold:true,color:'36566A',charSpacing:.3,margin:0});
slide.addText('КМТП  ·  Морской рыбный порт  ·  БГАРФ  ·  КМРК  ·  Правительство области  ·  Центр занятости',{x:.8,y:6.58,w:7.7,h:.2,fontSize:8.5,color:'78909E',margin:0,fit:'shrink'});
slide.addShape(pptx.ShapeType.roundRect,{x:11.15,y:6.32,w:.55,h:.48,rectRadius:.06,fill:{color:'A41E35'},line:{color:'A41E35'}}); slide.addText('ЦУР\n8',{x:11.15,y:6.42,w:.55,h:.2,fontSize:8,bold:true,color:'FFFFFF',align:'center',margin:0});
slide.addShape(pptx.ShapeType.roundRect,{x:11.82,y:6.32,w:.55,h:.48,rectRadius:.06,fill:{color:'19486A'},line:{color:'19486A'}}); slide.addText('ЦУР\n17',{x:11.82,y:6.42,w:.55,h:.2,fontSize:8,bold:true,color:'FFFFFF',align:'center',margin:0});
slide.addText('КАЛИНИНГРАД  ·  ПОРТ, КОТОРЫЙ РАСТИТ БУДУЩЕЕ',{x:8.35,y:7.08,w:4.2,h:.12,fontSize:6.5,color:'9AB0BA',charSpacing:.5,align:'right',margin:0});
// PowerPoint entrance transition; object motion can be added in PowerPoint Animation Pane.
slide.transition = { type: 'fade', duration: 0.7 };
pptx.writeFile({ fileName: 'molodezhny-eko-patrul-porta.pptx' });
