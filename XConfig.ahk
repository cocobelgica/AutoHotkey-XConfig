class XConfig
{
	
	
	__New(src) {
		ObjInsert(this, "_", [])
		ObjInsert(this, "__doc", ComObjCreate("MSXML2.DOMDocument.6.0"))
		this.async := false

		;Load XML source
		if (src ~= "s)^<.*>$")
			this.loadXML(src)
		else if ((f:=FileExist(src)) && !(f ~= "D"))
			this.load(src)
		else throw Exception("Invalid XML source.", -1)
		
	}

	__Set(k, v, p*) {

		if (k = "file")
			return this._[k] := v

		try if (n:=this.__doc.selectSingleNode(k)) {
			if ((nts:=n.nodeTypeString) = "element") {
				if ((c:=n.childNodes).length > 1)
					n := c.item(0)
				prev := n.text
				n.text := v

			} else if (nts = "attribute") {
				prev := n.nodeValue
				n.nodeValue := v
			
			}
			return prev
		}

		try return (this.__doc)[k] := v
	}

	class __Get extends XConfig.__PROPERTIES__
	{

		__(k, p*) {
			
			try if (n:=this.__doc.selectSingleNode(k)) {
				if ((nts:=n.nodeTypeString) = "element") {
					if (cc:=(c:=n.childNodes).length > 1)
						c := c.item(0)
					return p.1 ? n[p.1] : (cc ? c.text : n.text)

				} else if (nts = "attribute") {
					return n[p.1 ? p.1 : "nodeValue"]
				
				}
			
			}
			
			try return (this.__doc)[k]
		}

		file() {
			return this._.Haskey("file") ? this._.file : ""
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

	__Add(x, n, v*) {
		x := this.selectSingleNode(x)
		if IsObject(n) {
			for a, b in n
				x.setAttribute(a, b)
		
		} else if (n ~= "s)^<.*>$") {
			n := this.__Str2Node(n)
			, cmd := (r:=(p<>"")) ? "insertBefore" : "appendChild"
			, args := r ? [n, x.selectSingleNode(p)] : [n]
			
			return x[cmd](args*)

		} else {
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
		if ((nts:=(n:=this.selectSingleNode(x)).nodeTypeString) = "element")
			n.parentNode.removeChild(n)
		
		else if (nts = "attribute")
			p := this.selectSingleNode(RegExReplace(x, "/[^/]+$", ""))
			, p.removeAttributeNode(n)
	}

	__Save(dir:="", indent:=false) {

		if indent
			this.__Transform()

		this.save(dir<>"" ? dir : A_ScriptDir)
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

	__Str2Node(str) {
		static x

		if !x
			x := ComObjCreate("MSXML2.DOMDocument.6.0")
			, x.async := false

		x.loadXML("<ROOT>" str "</ROOT>")
		return x.documentElement.childNodes.item(0)
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