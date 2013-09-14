class XConfig
{
	
	
	__New(src, file:="") {
		ObjInsert(this, "_", []) ;Proxy object
		ObjInsert(this, "__dom", ComObjCreate(this.__MSXML()))
		this.setProperty("SelectionLanguage", "XPath") ;for OS<VISTA|7|8
		this.async := false

		;Load XML source
		if (src ~= "s)^<.*>$")
			this.loadXML(src)
		else if ((f:=FileExist(src)) && !(f ~= "D"))
			this.load(src)
		else throw Exception("Invalid XML source.", -1)

		if (file <> "")
			this.__file := file
	}

	__Set(k, v, p*) {

		if (k ~= "i)^__(file)$")
			return this._[k] := v

		try if (n:=this.__dom.selectSingleNode(k)) {
			if ((nts:=n.nodeTypeString) = "element") {
				if (t:=n.selectSingleNode("./text()")) {
					prev := t.nodeValue
					, t.nodeValue := v
				
				} else {
					prev := "" , t := this.createTextNode(v)
					if n.hasChildNodes()
						n.insertBefore(t, n.firstChild)
					
					else n.appendChild(t)
				}
			
			} else if (nts ~= "i)^(attribute|text|comment|cdatasection)$") {
				prev := n.nodeValue
				n.nodeValue := v
			
			}
			return prev
		}

		try return (this.__dom)[k] := v
	}

	class __Get extends XConfig.__PROPERTIES__
	{
		/*
		__(k, p*) {
			
			try if (n:=this.__dom.selectSingleNode(k)) {
				if ((nts:=n.nodeTypeString) = "element") {
					return p.1
					       ? n[p.1]
					       : ((t:=n.selectSingleNode("./text()")) ? t.nodeValue : "")

				} else if (nts ~= "i)^(attribute|text|comment|cdatasection)$") {
					return n[p.1 ? p.1 : "nodeValue"]
				
				}
			
			}
			
			try return (this.__dom)[k]
		}
		*/
		__(k, p*) {

			try if (n:=this.__dom.selectSingleNode(k)) {
				if p.MinIndex() {
					for a, b in p
						n := n[b]
					return n
				}
				
				if ((nts:=n.nodeTypeString) = "element")
					return ((t:=n.selectSingleNode("./text()")) ? t.nodeValue : "")

				else if (nts ~= "i)^(attribute|text|comment|cdatasection)$")
					return n.nodeValue
			}

			try return (this.__dom)[k]
		}

		__file() {
			return this._.Haskey("__file")
			       ? this._.__file
			       : ((url:=this.url)<>"" ? url : "")
		}

		__root() {
			return this.documentElement
		}
	}
	
	__Call(m, p*) {
		/*
		Do not initialize 'BIF' as class static initializer(s) will not be
		able to access the variable's content when calling this function.
		*/
		static BIF

		if !BIF
			BIF := "i)^(
			(LTrim Join|
			Insert
			Remove
			(Min|Max)Index
			(Set|Get)Capacity
			GetAddress
			_NewEnum
			HasKey
			Clone
			))$"

		if (!ObjHasKey(XConfig, m) && !(m~=BIF))
			try return (this.__dom)[m](p*)
	}

	__Add(x, n, p:="") {
		x := this.selectSingleNode(x)
		if IsObject(n) {
			for a, b in n
				x.setAttribute(a, b)
		
		} else if (n ~= "s)^<.*>$") {
			n := this.__XML2DOM(n)
			, cmd := (r:=(p<>"")) ? "insertBefore" : "appendChild"
			, args := r ? [n, x.selectSingleNode(p)] : [n]
			
			return x[cmd](args*) ; Fix this in case DocumentFragment is added.

		} else if (n ~= "i)^(?!(?:xml|[\d\W_]))[^\s\W]+$") { ; valid tagName
			e := this.createElement(n)
			if IsObject(p) {
				cmd := (r:=p.HasKey("ref")) ? "insertBefore" : "appendChild"
				, args := r ? [e, x.selectSingleNode(p.ref)] : [e]
				, e := x[cmd](args*)
				
				if p.HasKey("att")
					for a, b in p.att
						e.setAttribute(a, b)

				if p.HasKey("text")
					e.text := p.text
			
			} else {
				e := x.appendChild(e)
				if (p <> "")
					e.text := p
			}

			return e
		}
		return true
	}

	__Del(x) {
		
		if ((nts:=(n:=this.selectSingleNode(x)).nodeTypeString) = "attribute") {
			/*
			;This does not work on XP and below
			for e in this.selectNodes("//*[@" n.name "='" n.value "']")
				continue
			*/
			Loop % (_:=this.selectNodes("//*[@" n.name "='" n.value "']")).length
				e := _.item(A_Index-1)
			until e.selectNodes("@*").matches(n)
			e.removeAttributeNode(n)
		
		} else if (nts ~= "i)^(element|text|comment|cdatasection)$")
			n.parentNode.removeChild(n)
	}

	__Save(dir:="", indent:=false) {

		if indent
			this.__Transform()

		this.save(dir<>""
		         ? dir
		         : ((f:=this.__file) ? f : A_WorkingDir "\XCONFIG-" A_TickCount))
	}

	__Transform() {
		static xsl

		if !xsl {
			xsl := ComObjCreate(this.__MSXML())
			style := "
			(LTrim
			<xsl:stylesheet version='1.0' xmlns:xsl='http://www.w3.org/1999/XSL/Transform'>
			<xsl:output method='xml' indent='yes' encoding='UTF-8'/>
			<xsl:template match='@*|node()'>
			<xsl:copy>
			<xsl:apply-templates select='@*|node()'/>
			<xsl:for-each select='@*'>
			<xsl:text></xsl:text>
			</xsl:for-each>
			</xsl:copy>
			</xsl:template>
			</xsl:stylesheet>
			)"
			xsl.loadXML(style)
		}
		this.transformNodeToObject(xsl, this.__dom)
	}
	
	__XML2DOM(str) {
		static x

		if !x
			x := ComObjCreate(this.__MSXML())
			, x.async := false

		x.loadXML("<XCONFIG>" str "</XCONFIG>")
		n := this.ownerDocument.importNode(x.documentElement, true)
		DOMNode := (n.childNodes.length>1)
		        ? this.createDocumentFragment()
		        : n.removeChild(n.firstChild)

		while (n.hasChildNodes())
			DOMNode.appendChild(n.removeChild(n.firstChild))
		
		return DOMNode
	}
	/*
	Short-hand for selectNodes/selectSingleNode
	*/
	__(xpr, single:=true) {
		;Bypass __Call in this case
		return (this.__dom)[single ? "selectSingleNode" : "selectNodes"](xpr)
	}

	__Sel(xpr) {
		return new XConfig.__NODE__(this.__(xpr))
	}
	/*
	Returns the node type of a node represented as XML string.
	*/
	__Type(str, string:=true) {
		static r

		if !r
			r := {a:{0:2, 1:"attribute"}
		        , cds:{0:4, 1:"cdatasection"}
		        , c:{0:8, 1:"comment"}
		        , e:{0:1, 1:"element"}}

		;attribute
		if (str ~= "^[\w]+=(""|')(?:(?!\1).)*?\1$")
			return r["a", string]
		;cdatasection
		else if (str ~= "s)^<!\[CDATA\[(?:(?!]]>).)*?]]>$")
			return r["cds", string]
		;comment
		else if (str ~= "s)^<!--.*?-->$")
			return r["c", string]
		;element
		else if (str ~= "s)^<((?!(?:(?i)xml|[\d\W_]))[^\s\W]+)[^>]*?(?:/>$|>.*?</\1\s*>)$")
			return r["e", string]

		else throw Exception("No match", -1)
	}

	__MSXML() {
		static MSXML := XConfig.__MSXML()

		if !MSXML
			MSXML := "MSXML2.DOMDocument"
			      . ((A_OsVersion~="^WIN_(VISTA|7|8)$") ? ".6.0" : "")

		return MSXML
	}
	/*
	Private Method
	__RGX(type:="element") {
		static xpr , k

		if !xpr {
			xpr := "
			(LTrim
			^[\w]+=(""|')(?:(?!\1).)*?\1$
			s)^<!\[CDATA\[(?:(?!]]>).)*?]]>$
			s)^<!--.*?-->$
			s)^<((?!(?:(?i)xml|[\d\W_]))[^\s\W]+)(?:[^>]+|)(?:/>$|>.*?</\1\s*>)$
			i)^(?!(?:xml|[\d\W_]))[^\s\W]+$
			)"
			k := {attribute:1,cdatasection:2,comment:3,element:4,tagName:5}
		}
		;RegExMatch(xpr, "(?:[^\r\n]+\R){" k[type]-1 "}\K[^\r\n]+", m)
		RegExMatch(xpr, "(?:\R?\K[^\r\n]+){" k[type] "}", m)
		return m
	}
	*/
	class __NODE__
	{

		__New(oContext, n:=".") {
			ObjInsert(this, "_", [])
			ObjInsert(this, "__dom", IsObject(n) ? n : oContext.selectSingleNode(n))
		
		}

		__Set(k, v, p*) {

			if (n:=this.__(k))
				return n[(n.nodeType>1 ? "nodeValue" : "text")] := v
			
			else if (k ~= "i)^@\w+$")
				return this.setAttribute(SubStr(k, 2), v)
		}

		class __Get extends XConfig.__PROPERTIES__
		{

			__(k, p*) {
				static DOMNode_Property

				if !DOMNode_Property
					DOMNode_Property := "i)^(
					(LTrim Join|
					attributes
					baseName
					childNodes
					dataType
					definition
					(first|last)Child
					namespaceURI
					(next|previous)Sibling
					node(Name|Type(dValue|String)?|Value)
					ownerDocument
					parentNode
					parsed
					prefix
					specified
					tagName
					text
					xml
					))$"

				if (k ~= DOMNode_Property)
					try return (this.__dom)[k]

				else return XConfig.__Get.__.(this, k, p*)
			}
		}

		__Call(m, p*) {
			static DOMNode_Method

			if !DOMNode_Method
				DOMNode_Method := "i)^(
				(LTrim Join|
				(append|remove|replace)Child
				cloneNode
				get(Attribute(Node)?|ElementsByTagName)
				hasChildNodes
				insertBefore
				normalize
				removeAttribute(Node)?
				select(Nodes|SingleNode)
				setAttribute(Node)?
				transformNode(ToObject)?
				))$"
			
			if ObjHasKey(XConfig, m)
				return (m<>"__Add")
				       ? XConfig[m].(this, p*)
				       : XConfig[m].(this, ".", p*)

			else if (m ~= DOMNode_Method)
				return XConfig.__Call.(this, m, p*)
		}
	
	}

	class __PROPERTIES__
	{

		__Call(target, name, params*) {
			if !(name ~= "i)^(base|__Class)$") {
				return ObjHasKey(this, name)
				       ? this[name].(target, params*)
				       : this.__.(target, name, params*)
			}
		}
	}
}