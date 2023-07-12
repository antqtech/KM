  <!--網站連結--->
  <a href="https://www.w3schools.com">This is a link</a>
  
  <!--src圖片  alt圖片屬性--->  
  <img src="w3schools.jpg" alt="W3Schools.com" width="104" height="142">
  <title>Document</title>
  
  <!--HTML Elements--> 
  <p>This is a <br> paragraph with a line break.</p>
  
  <!--HTML Heading--> 
  <h2 style="font-size:60px;" title="I'm a header">The title Attribute</h2>
  <p title="I'm a tooltip">Mouse over this paragraph, to display the title attribute as a tooltip.</p>
 
  <!--HTML Paragraphs--> 
  <h1>This is heading 1</h1>
  <p>This is some text.</p>
  <hr>//一條分隔線
  <br>//類似enter下一段
  <pre> element defines preformatted text.

  <!---style--->
  <h1 style="background-color:powderblue;">This is a heading</h1>//字體背景顏色
  <h1 style="color:blue;">This is a heading</h1>//字體顏色 
  <h1 style="font-family:verdana;">This is a heading</h1>
  <h1 style="font-size:300%;">This is a heading</h1>//大小
  <h1 style="text-align:center;">Centered Heading</h1>//位置
 
  <!--HTML Text Formatting-->
    <b> - Bold text //粗
    <strong> - Important text //粗
    <i> - Italic text  //斜體
    <em> - Emphasized text //斜體
    <mark> - Marked text //黃色螢光筆
    <small> - Smaller text //變小
    <del> - Deleted text //刪除的線線</del>
    <ins> - Inserted text //底線
    <sub> - Subscript text //下標
    <sup> - Superscript text //上標
  <!---HTML Quotation and Citation Elements--->
  <blockquote> element defines a section that is quoted from another source.//文字縮排
  <q>//Quotations“+”
  <abbr> Abbreviations//標記縮寫的
  <address> //for Contact Information
  <cite> for Work Title//(e.g. a book, a poem, a song, a movie, a painting, a sculpture, etc.).
    <bdo> for Bi-Directional Override//字的排列相反


  <!--HTML Comments-->
  <!--HTML Colors-->
  <p style="color:DodgerBlue;">Lorem ipsum...</p>
  <h1 style="border:2px solid Tomato;">Hello World</h1>
  <h1 style="background-color:DodgerBlue;">Hello World</h1>
  //顏色條配的網站https://htmlcolorcodes.com/


  <!--HTML Styles - CSS  尚未讀完--->
  <!---HTML Links--->
  <a href="https://www.w3schools.com/" target="_blank">Visit W3Schools!</a>//The target Attribute
  <a href="https://www.w3schools.com/">Visit W3Schools.com!</a>
  <a href="default.asp">
  <img src="smiley.gif" alt="HTML tutorial" style="width:42px;height:42px;">
  </a>//Use an Image as a Link
  <a href="mailto:someone@example.com">Send email</a>//Link to an Email Address  
  <button onclick="document.location='default.asp'">HTML Tutorial</button>//Button as a Link
  
  <!--HTML Images-->
  <img src="img_chania.jpg" alt="Flowers in Chania" style="width:128px;height:128px;">
  
  //會浮動的
  <p><img src="smiley.gif" alt="Smiley face" style="float:right;width:42px;height:42px;">
    The image will float to the right of the text.</p>
    
    <p><img src="smiley.gif" alt="Smiley face" style="float:left;width:42px;height:42px;">
    The image will float to the left of the text.</p>

  <!--map-->
  <img src="workplace.jpg" alt="Workplace" usemap="#workmap" width="400" height="379">
  <map name="workmap">
  <area shape="rect" coords="34,44,270,350" alt="Computer" href="computer.htm">
  <area shape="rect" coords="290,172,333,250" alt="Phone" href="phone.htm">
  <area shape="circle" coords="337,300,44" alt="Cup of coffee" href="coffee.htm">
  </map>


 <!---HTML Favicon--->
 用在網站名稱旁邊的小圖示icon
  <head>
  <title>My Page Title</title>
  <link rel="icon" type="image/x-icon" href="/images/favicon.ico">
</head>

 <!---HTML table--->
<table>
  <tr>
    <th>Company</th>
    <th>Contact</th>
    <th>Country</th>
  </tr>
  <tr>
    <td>Alfreds Futterkiste</td>
    <td>Maria Anders</td>
    <td>Germany</td>
  </tr>
  <tr>
    <td>Centro comercial Moctezuma</td>
    <td>Francisco Chang</td>
    <td>Mexico</td>
  </tr>
</table>