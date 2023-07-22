// 圖片列表
const imageList = [
    'url("image/1.jpg")',
    'url("image/2.jpg")',
    'url("image/3.jpg")',
    'url("image/4.jpg")',
    // 添加更多的URL
  ];
  
  // 获取background元素
const background = document.querySelector('.background');

// 定义变量来追踪当前显示的图片索引
let currentIndex = 0;

// 定义动画函数
function moveBackgroundAnimation() {
  background.style.backgroundImage = imageList[currentIndex];

  currentIndex = (currentIndex + 1) % imageList.length;

  setTimeout(moveBackgroundAnimation, 10000); // 延时15秒后再次调用动画函数
}

// 开始动画
moveBackgroundAnimation();