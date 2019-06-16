'Примітка: * - Microsoft розширення W3C DOM
Public Sub main()
Dim xmldoc As New MSXML2.DOMDocument50 'створити об'єкт xml документ
'або так:
'Dim xmldoc As MSXML2.DOMDocument50
'Set xmldoc = CreateObject("Msxml2.DOMDocument.5.0")
Dim docNode, el1Node, el2Node, node As MSXML2.IXMLDOMNode 'об'єкти xml-вузли
Dim txtNode As MSXML2.IXMLDOMText 'об'єкт xml текстовий вузол
Dim el As MSXML2.IXMLDOMElement 'об'єкт xml-елемент
Dim atr As MSXML2.IXMLDOMAttribute 'об'єкт xml-атрибут

'*завантажити документ з рядка
xmldoc.loadXML _
        "<?xml version='1.0'?>" + vbNewLine + _
        "<doc title='my'>" + vbNewLine + _
        "   <el1 atr1='1'>" + vbNewLine + _
        "      <el2 atr1='hello'>" + vbNewLine + _
        "          Hello world!" + vbNewLine + _
        "      </el2>" + vbNewLine + _
        "   </el1>" + vbNewLine + _
        "</doc>" + vbNewLine
xmldoc.Save "e:\\my.xml" '*зберегти документ

Dim b As Boolean
b = xmldoc.Load("e:\\my.xml") '*завантажити документ
If b Then 'якщо завантажено, тоді
    Set docNode = xmldoc.documentElement 'кореневий елемент
    Set el1Node = docNode.FirstChild 'перший дочірній вузол вузла docNode
    Set el2Node = el1Node.FirstChild 'перший дочірній вузол вузла el1Node
    Set txtNode = el2Node.FirstChild 'перший дочірній вузол вузла el2Node
    Debug.Print txtNode.NodeValue 'вивести текстове значення вузла
    
    Set el = xmldoc.createElement("el3") 'створити вузол елемента з іменем el3
    el.Text = "hello!!!" '*текстовий вміст вузла і підвузлів
    Set atr = xmldoc.createAttribute("attr") 'створити атрибут з іменем attr
    atr.Value = "10" 'значення атрибута
    el.setAttributeNode atr 'установити вузол атрибута atr в елемент el
    docNode.appendChild el 'додати дочірній вузол як останній
    
    Set node = docNode.LastChild.CloneNode(True) 'клонувати останній дочірній вузол з підвузлами
    docNode.InsertBefore node, el1Node 'вставити дочірній вузол перед el1Node
    docNode.RemoveChild node 'видалити дочірній вузол node
    
    docNode.appendChild xmldoc.createTextNode("hello") 'додати дочірній текстовий вузол
    
    Dim s As Variant
    Set node = docNode.ChildNodes.Item(0) 'дочірній вузол з індексом 0 (перший)
    s = node.nodeName 'ім'я вузла
    s = node.NodeType 'тип вузла
    s = node.NodeValue 'текстове значення вузла
    s = node.HasChildNodes 'чи вузол має дочірні вузли
    s = node.Text '*текстовий вміст вузла і підвузлів
    s = node.XML '*XML вузла і підвузлів
    s = node.DataType '*тип даних вузла
    s = node.parsed '*перевіряє чи вузол і підвузли проаналізовані
    s = node.ParentNode.nodeName 'ім'я батьківського вузла
    s = node.NextSibling.nodeName 'ім'я наступного спорідненого вузла
    s = node.ChildNodes.Item(0).nodeName 'ім'я дочірнього вузла з індексом 0
    s = node.Attributes.Length 'кількість атрибутів вузла
    s = node.Attributes.Item(0).NodeValue 'текстове значення атрибута з індексом 0
    s = node.Attributes.Item(0).specified '*чи заданий явно, чи за замовчуванням
    s = docNode.ChildNodes.Item(0).OwnerDocument.nodeName 'ім'я кореня документа
    
    s = xmldoc.getElementsByTagName("el2").Length 'кількість елементів з тегом el2
    Set el = xmldoc.getElementsByTagName("el2").Item(0) 'перший елемент з тегом el2
    s = el.nodeName 'ім'я вузла
    s = el.tagName 'ім'я тега
    s = el.GetAttribute("atr1") 'значення атрибута atr1
    el.setAttributeNode xmldoc.createAttribute("Data") 'створити і установити атрибут Data
    el.setAttribute "Data", "today" 'задати значення атрибуту Data
    s = el.getAttributeNode("Data").Value 'значення атрибута Data
    Debug.Print s
    
    xmldoc.Save "e:\\my.xml" '*зберегти документ
End If
End Sub
