
export default function createPopover(id, node, children) {
  let popoverContainer = document.createElement("div");
  popoverContainer.id = `popover-comment-${id}`;
  popoverContainer.className = "popover-container";
  popoverContainer.style.padding = "15px";
  popoverContainer.style.boxShadow = "3px 3px 8px #ccc";
  popoverContainer.style.position = "fixed";
  popoverContainer.style.zIndex = "999";
  popoverContainer.style.backgroundColor = "#fff";
  popoverContainer.style.borderRadius = "5px";

  popoverContainer.appendChild(children);

  document.body.appendChild(popoverContainer);
  
  let rect = node.getBoundingClientRect();
  popoverContainer.style.top = `${rect.top - 20}px`;
  popoverContainer.style.left = `${rect.left + 35}px`;

  var isVisible = true;

  function hidePopover() {
    isVisible = false;
    document.body.removeChild(popoverContainer);
  }

  document.addEventListener('click', function (e) {
    e.stopPropagation();
    if (isVisible &&
        e.target !== popoverContainer &&
        e.target.parentElement !== popoverContainer &&
        e.target !== node &&
        e.target.parentElement !== node
    ) {
      hidePopover();
    }
  });

  popoverContainer.addEventListener('mouseleave', function (e) {
    hidePopover();
  })
}
