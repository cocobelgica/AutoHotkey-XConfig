class XConfig
{
	
	
	__New(src, file:="") {
		ObjInsert(this, "_", []) ;Proxy object
		ObjInsert(this, "__doc", ComObjCreate("MSXML2.DOMDocument.6.0"))
		this.setProperty("SelectionLanguage", "XPath") ; Not really needed.
		this.async := false

		;Load XML source
		if (src ~= "s)^<.*>$")
			this.loadXML(src)
		else if ((f:=FileExist(src)) && !(f ~= "D"))
			this.load(src)
		else throw Exception("Invalid XML source.", -1)

		if (file <> "")
			this.file := file
	}

	__Set(k, v, p*) {

		if (k = "file")
			return this._[k] := v

		try if (n:=this.__doc.selectSingleNode(k)) {
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

		try return (this.__doc)[k] := v
	}

	class __Get extends XConfig.__PROPERTIES__
	{
		/*
		__(k, p*) {
			
			try if (n:=this.__doc.selectSingleNode(k)) {
				if ((nts:=n.nodeTypeString) = "element") {
					return p.1
					       ? n[p.1]
					       : ((t:=n.selectSingleNode("./text()")) ? t.nodeValue : "")

				} else if (nts ~= "i)^(attribute|text|comment|cdatasection)$") {
					return n[p.1 ? p.1 : "nodeValue"]
				
				}
			
			}
			
			try return (this.__doc)[k]
		}
		*/
		__(k, p*) {

			try if (n:=this.__doc.selectSingleNode(k)) {
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

			try return (this.__doc)[k]
		}

		file() {
			return this._.Haskey("file")
			       ? this._.file
			       : ((url:=this.url)<>"" ? url : "")
		}

		root() {
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

		if (!ObjHasKey(this.base, m) && !(m ~= BIF))
			try return (this.__doc)[m](p*)
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
			for e in this.selectNodes("//*[@" n.name "='" n.value "']")
				continue
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
		         : ((f:=this.file) ? f : A_WorkingDir "\XCONFIG-" A_TickCount))
	}

	__Transform() {
		static xsl

		if !xsl {
			xsl := ComObjCreate("MSXML2.DOMDocument.6.0")
			style := "
			(LTrim
			<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">
			<xsl:output method=""xml"" indent=""yes"" encoding=""UTF-8""/>
			<xsl:template match=""@*|node()"">
			<xsl:copy>
			<xsl:apply-templates select=""@*|node()""/>
			<xsl:for-each select=""@*"">
			<xsl:text></xsl:text>
			</xsl:for-each>
			</xsl:copy>
			</xsl:template>
			</xsl:stylesheet>
			)"
			xsl.loadXML(style)
		}
		this.transformNodeToObject(xsl, this.__doc)
	}
	/*
	__Str2Node(str) {
		static x

		if !x
			x := ComObjCreate("MSXML2.DOMDocument.6.0")
			, x.async := false

		x.loadXML("<XCONFIG>" str "</XCONFIG>")

		if (pe:=x.parseError).errorCode {
			RegExMatch(str, "sO)^<([^\s>]+)(?:[^>]+|)>", t)
			, RegExMatch(str, "sO)^" t.value "(?:.*?|)(</" t.1 ">|$)$", m)
			
			if (m.1="") && (pe.reason~="i)(end|start|tag|not|match|XCONFIG|" t.1 ")")
				return this.__Str2Node(m.value . "</" t.1 ">")
			
			else throw Exception(pe.reason, -1)
		;} else return x.documentElement.firstChild
		} else return this.importNode(x.documentElement.firstChild, true)
	}
	*/
	__XML2DOM(str) {
		static x

		if !x
			x := ComObjCreate("MSXML2.DOMDocument.6.0")
			, x.async := false

		x.loadXML("<XCONFIG>" str "</XCONFIG>")
		n := this.importNode(x.documentElement, true)
		DOMNode := (n.childNodes.length > 1)
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
		return (this.__doc)[single ? "selectSingleNode" : "selectNodes"](xpr)
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
		else if (str ~= "s)^<((?!(?:(?i)xml|[\d\W_]))[^\s\W]+)(?:[^>]+|)(?:/>$|>.*?</\1\s*>)$")
			return r["e", string]

		else throw Exception("No match", -1)
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