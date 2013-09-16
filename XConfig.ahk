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
					prev := "" , t := this.__doc.createTextNode(v)
					n.hasChildNodes()
					? n.insertBefore(t, n.firstChild)
					: n.appendChild(t)
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

		__doc() {
			return this.__dom
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
			for k, v in n {
				if RegExMatch(k, "Oi)^@(\w*)$", att)
					if (att[1] <> "")
						x.setAttribute(att[1], v)

					else for a, b in v
						x.setAttribute(a, b)

				else if (k ~= "i)^t(ext(\(\))?)?$")
					x.text := v

				else if (k ~= "i)^cd(s|ata(section)?)$")
					x.hasChildNodes()
					? x.insertBefore(this.__doc.createCDATASection(v), x.firstChild)
					: x.appendChild(this.__doc.createCDATASection(v))
			}
		
		} else if (n ~= "s)^<.*>$") {
			n := this.__XML2DOM(n)
			, cmd := (r:=(p<>"")) ? "insertBefore" : "appendChild"
			, args := r ? [n, x.selectSingleNode(p)] : [n]
			
			return x[cmd](args*) ; Fix this in case DocumentFragment is added.

		} else if (n ~= "i)^(?!(?:xml|[\d\W_]))[^\s\W]+$") { ; valid tagName
			e := this.__doc.createElement(n)
			if IsObject(p) {
				cmd := (r:=p.Remove("ref")) ? "insertBefore" : "appendChild"
				, args := (r<>"") ? [e, x.selectSingleNode(r)] : [e]
				, e := new XConfig.__NODE__(x[cmd](args*))
				, e.__Add(p)
			
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

	__Transform(DOM:=false) {
		static xsl

		if !xsl {
			xsl := ComObjCreate(this.__MSXML())
			style := "
			(LTrim Join
			<?xml version='1.0' encoding='ISO-8859-15'?>
			<xsl:stylesheet xmlns:xsl='http://www.w3.org/1999/XSL/Transform' version='1.0'>
			<xsl:output method='xml'/>

			<xsl:template match='@*'>
			<xsl:copy/>
			</xsl:template>

			<xsl:template match='text()'>
			<xsl:value-of select='normalize-space(.)' />
			</xsl:template>

			<xsl:template match='*'>
			<xsl:param name='indent' select='""""'/>
			<xsl:text>&#xa;</xsl:text>
			<xsl:value-of select='$indent' />
			<xsl:copy>
			<!-- <xsl:apply-templates select='@*|*|text()'> -->
			<xsl:apply-templates select='@*|node()'>
			<xsl:with-param name='indent' select='concat($indent, ""  "")'/>
			</xsl:apply-templates>
			</xsl:copy>
			<xsl:if test='count(../*)>0 and ../*[last()]=.'>
			<xsl:text>&#xa;</xsl:text>
			<xsl:value-of select='substring($indent,3)' />
			</xsl:if>
			</xsl:template>

			</xsl:stylesheet>
			)"
			xsl.loadXML(style)
		}
		if DOM
			this.transformNodeToObject(xsl, IsObject(DOM) ? DOM : this.__doc)
		else return this.transfromNode(xsl)
	}
	/*
	Converts a node[element] represented as an XML string to DOM object
	*/
	__XML2DOM(str) {
		static x

		if !x
			x := ComObjCreate(this.__MSXML())
			, x.async := false

		x.loadXML("<XCONFIG>" str "</XCONFIG>")
		n := this.__doc.importNode(x.documentElement, true)
		DOMNode := (n.childNodes.length>1)
		        ? this.__doc.createDocumentFragment()
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
	/*
	Works like selectSingleNode but returns an IXMLDOMNode object
	wrapped/subclassed as an XConfig.__NODE__ object.
	*/
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

			__doc() {
				return this.ownerDocument
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
				;return XConfig.__Call.(this, m, p*)
				try return (this.__dom)[m](p*)
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