export const createIconPinglun = function() {
  let svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 1024 1024");
  svg.setAttribute("width", '20px');
  svg.setAttribute("height", '20px');
  let path = document.createElementNS("http://www.w3.org/2000/svg", "path");
  path.setAttribute("d", "M512.13274 64C229.33274 64 0.13274 250.2 0.13274 480c0 95.2 39.8 182.4 105.8 252.6C76.13274 811.4 14.13274 878.2 13.13274 879c-13.2 14-16.8 34.4-9.2 52S28.93274 960 48.13274 960c123 0 220-51.4 278.2-92.6C384.13274 885.6 446.53274 896 512.13274 896c282.8 0 512-186.2 512-416S794.93274 64 512.13274 64z m0 736c-53.4 0-106.2-8.2-156.8-24.2l-45.4-14.4-39 27.6c-28.6 20.2-67.8 42.8-115 58 14.6-24.2 28.8-51.4 39.8-80.4l21.2-56.2-41.2-43.6C139.53274 628.2 96.13274 564.4 96.13274 480c0-176.4 186.6-320 416-320s416 143.6 416 320-186.6 320-416 320z");
  path.setAttribute("fill", "#777");
  svg.appendChild(path);
  return svg;
}

export const createIconCollapse = function() {
  let svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 1024 1024");
  svg.setAttribute("width", '24px');
  svg.setAttribute("height", '24px');
  let path = document.createElementNS("http://www.w3.org/2000/svg", "path");
  path.setAttribute("d", "M688.553 587.447L383.686 851.786c-20.841 18.07-52.391 15.83-70.468-5.003A49.927 49.927 0 0 1 301 814.063V209.936c0-27.58 22.365-49.937 49.955-49.937a49.965 49.965 0 0 1 32.731 12.213l304.867 264.34c41.683 36.141 46.165 99.219 10.01 140.887a99.891 99.891 0 0 1-10.01 10.007z");
  path.setAttribute("fill", "#fff");
  svg.appendChild(path);
  return svg;
}

export const createIconPrev = function() {
  let svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 1024 1024");
  svg.setAttribute("width", '30px');
  svg.setAttribute("height", '30px');
  let path = document.createElementNS("http://www.w3.org/2000/svg", "path");
  path.setAttribute("d", "M640 514.24L445.76 320l-48 48L544 514.24l-146.24 146.24 48 48L640 514.24z");
  path.setAttribute("fill", "#999");
  svg.appendChild(path);
  return svg;
}
