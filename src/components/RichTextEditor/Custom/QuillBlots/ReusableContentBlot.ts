
import { Quill } from 'react-quill';
import { toBoolean, isset } from '@spfxappdev/utility';
const BlockEmbed = Quill.import('blots/block/embed');


export class ReusableContentBlot extends BlockEmbed {
    public static blotName = "reusable";
    public static tagName = "div";
    public static className = `ql-custom`;

    public static create(data) {
        const node = super.create();
        const denotationChar = document.createElement("span");
        denotationChar.className = "ql-reusable";
        denotationChar.innerHTML = data.value;
        
        if(isset(data.isStatic) && toBoolean(data.isStatic) == true) {
            denotationChar.setAttribute("contenteditable", "false");
        }
        
        node.appendChild(denotationChar);
        //node.innerHTML += data.value;
        return ReusableContentBlot.setDataValues(node, data);
    }
  
    public static setDataValues(element, data) {
      const domNode = element;
      Object.keys(data).forEach(key => {
        domNode.dataset[key] = data[key];
      });
      return domNode;
    }
  
    public static value(domNode) {
      return domNode.dataset;
    }
}