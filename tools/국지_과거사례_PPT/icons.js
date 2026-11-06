const React=require('react'); const {renderToStaticMarkup}=require('react-dom/server'); const sharp=require('sharp');
const g=require('react-icons/gi');
const list=process.argv.slice(2); // name:color
(async()=>{for(const it of list){const [n,c]=it.split(':'); const Ico=g[n]; if(!Ico){console.log('missing',n);continue;}
 const svg=renderToStaticMarkup(React.createElement(Ico,{size:512,color:'#'+c}));
 await sharp(Buffer.from(svg)).png().toFile(`${__dirname}/icons/${n}_${c}.png`);} console.log('done');})();
