import Base from "../document"
import OfficeDocument from "./officeDocument"

export default class extends Base{
	static get ext(){return 'docx'}

	static OfficeDocument=OfficeDocument

	isProperty(node){
		let {name,parent}=node
		let tag=name.split(':').pop()
		if(super.isProperty(...arguments) || tag=='tblGrid')
			return true

		if(parent && parent.name && parent.name.split(':').pop()=='inline')
			return true

		return false
	}

	createElement(node){
		let {name, attributes:{directStyle}}=node
		let type=name.split(':').pop()
		switch(type){
		case "p":
			type="paragraph"
		break
		case "r":
			type="inline"
		break
		case "t":
			type="text"
		break
		case "tbl":
			type="table"
		break
		case "tr":
			type="row"
		break
		case "tc":
			type="cell"
		break
		case "hdr":
			type="header"
		break
		case "ftr":
			type="footer"
		break
		case "inline":
			break
		case "drawing":
			break
		case "sdt":
			break
		}

		return this.onCreateElement(node, type)
	}

	toProperty(node, type){
		return {};
	}

	onToControlProperty(node,type){
		switch(type){
		case 'dataBinding':
			let key=node.$.xpath.split(/[\/\:\[]/).splice(-2,1)
			node.parent.$.control={type:'documentProperty', key}
		break
		case 'text':
			if(!node.parent.$.control)
				node.parent.$.control={type:`control.${type}`}
		break
		case 'picture':
		case 'docPartList':
		case 'comboBox':
		case 'dropDownList':
		case 'date':
		case 'checkbox':
			node.parent.$.control={type:`control.${type}`}
		break
		case 'richtext':
			node.parent.$.control={type:"control.richtext"}
		break
		}
		return super.onToProperty(...arguments)
	}

	onToProperty(node, type){
		const {$:x, parent}=node
		if(parent && parent.name=='w:sdtPr')
			return onToControlProperty(...arguments)
		let value
		switch(type){
		default:
			return super.onToProperty(...arguments)
		}
	}

	asToggle(x){
		if(x==undefined || x.val==undefined){
			return -1
		}else{
			return parseInt(x.val)
		}
	}

	toSpacing(x){
		var r=x, o={}

		if(!r.beforeAutospacing && r.beforeLines)
			o.top=this.dxa2Px((r.beforeLines))
		else if(r.before)
			o.top=this.dxa2Px((r.before))

		if(!r.afterAutospacing && r.afterLines)
			o.bottom=this.dxa2Px((r.afterLines))
		else if(r.after)
			o.bottom=this.dxa2Px((r.after))

		if(!r.line)
			return o

		switch(x.lineRule){
		case 'atLeast':
		case 'exact':
			o.lineHeight=this.dxa2Px((x.line))
			break
		case 'auto':
		default:
			o.lineHeight=(parseInt(r.line)*100/240)+'%'
		}
		o.lineRule=x.lineRule
		return o
	}

	toBorder(x){
		var border=x
		border.sz && (border.sz=this.pt2Px(border.sz/8));
		border.color && (border.color=this.asColor(border.color))
		return border
	}

	toHeaderFooter(node,tag){
		const {$:{id, type}}=node
		let part=new HeaderFooter(this.officeDocument.rels[id].target, this, type)
		return part.parse()
	}
}
