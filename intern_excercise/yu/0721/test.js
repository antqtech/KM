// 图片列表，替换为你自己的图片URL
const imageList = [
    // 'url(image/1.jpg)',
    'url(image/2.jpg)',
    'url(image/3.jpg)',
    // 添加更多的图片URL
  ];
  
  // 获取background元素
  const background = document.querySelector('.background');
  
  // 定义变量来追踪当前显示的图片索引
  let currentIndex = 0;
  
  // 定义动画函数
  function moveBackgroundAnimation() {
    // 为背景元素设置多重背景图像
    const backgroundImage = imageList.map((url) => {
      return url;
    }).join(',');
  
    background.style.backgroundImage = backgroundImage;
  
    currentIndex = (currentIndex + 1) % imageList.length;
  
    setTimeout(moveBackgroundAnimation, 10000); // 延时15秒后再次调用动画函数
  }
  
  // 开始动画
  moveBackgroundAnimation();
  