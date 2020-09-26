function Webbie(element) {
	var i, s;
	s = "";

	if (element.nodeType === 1) {
		// Element node, e.g. A or P
		if (element.nodeName === "A") {
			s = s + "Link: ";
		} else {
			// Handle without special casing.
		}
	}
	if (element.childNodes.length > 0) {
		for (i = 0; i < element.childNodes.length; i++)
		{
			s = s + Webbie(element.childNodes[i]);
		}
		return s;
	} else {
		if (element.nodeType === 3) {
			// Text node: return text content.
			return element.nodeValue;
		} else {
			// Other node: return nothing (e.g. comment)
			return "";
		}
	}
}