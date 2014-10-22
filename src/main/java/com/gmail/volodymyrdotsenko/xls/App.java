package com.gmail.volodymyrdotsenko.xls;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathFactory;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class App {
	public static boolean copyFile(File newFile, File templeteFile) {
		try {
			FileUtils.copyFile(templeteFile, newFile, false);

			return true;
		} catch (IOException e) {
			e.printStackTrace();
		}

		return false;
	}

	public static boolean zip(File file, File xmlFile) throws IOException {
		InputStream in = null;
		ZipInputStream zipIn = null;
		ZipOutputStream zipOut = new ZipOutputStream(new FileOutputStream(
				new File("NewWhiteBoard_mod.xlsm")));

		try {
			zipIn = new ZipInputStream(new FileInputStream(file));
			ZipEntry entry = zipIn.getNextEntry();
			while (entry != null) {

				ZipEntry e = new ZipEntry(entry.getName());

				if ("xl/worksheets/sheet2.xml".equals(entry.getName())) {

					zipOut.putNextEntry(e);

					in = new FileInputStream(xmlFile);

					IOUtils.copy(in, zipOut);
				} else {
					zipOut.putNextEntry(e);

					IOUtils.copy(zipIn, zipOut);
				}
				entry = zipIn.getNextEntry();
			}

			return true;
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (zipIn != null)
				zipIn.close();

			if (zipOut != null)
				zipOut.close();

			IOUtils.closeQuietly(in);
		}

		return false;
	}

	public static void main(String[] args) throws Exception {

		File newFile = new File("NewWhiteBoard.xlsm");

		if (!copyFile(newFile, new File("D:/NewWhiteBoard_v3.xlsm"))) {
			return;
		}

		OPCPackage pkg = OPCPackage.open(newFile);
		XSSFReader r = new XSSFReader(pkg);
		// SharedStringsTable sst = r.getSharedStringsTable();

		// XMLReader parser = new App().fetchSheetParser(sst);

		// rId2 found by processing the Workbook
		// Seems to either be rId# or rSheet#
		InputStream sheet2 = r.getSheet("rId2");

		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
		Document doc = dBuilder.parse(sheet2);

		doc.getDocumentElement().normalize();

		XPathFactory xpf = XPathFactory.newInstance();
		XPath xpath = xpf.newXPath();
		XPathExpression expression = xpath.compile("//sheetData");

		Node sheetData = (Node) expression.evaluate(doc, XPathConstants.NODE);

		NodeList nodes = sheetData.getChildNodes();

		boolean start = false;

		int num = nodes.getLength();

		for (int i = 0; i < num; i++) {
			Node n = nodes.item(i);

			int row = Integer.valueOf(n.getAttributes().getNamedItem("r")
					.getNodeValue());

			if (row > 21) {
				start = true;
			}

			if (start) {
				sheetData.removeChild(n);
				i--;
				num--;
			}
		}

		TransformerFactory tf = TransformerFactory.newInstance();
		Transformer transformer = tf.newTransformer();
		transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "no");
		transformer.setOutputProperty(OutputKeys.METHOD, "xml");
		// transformer.setOutputProperty(OutputKeys.INDENT, "yes");
		transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
		// transformer.setOutputProperty(
		// "{http://xml.apache.org/xslt}indent-amount", "4");

		File xmlFile = new File("sheet2.xml");
		FileOutputStream out = new FileOutputStream(xmlFile);
		transformer.transform(new DOMSource(doc), new StreamResult(
				new OutputStreamWriter(out, "UTF-8")));

		sheet2.close();

		zip(newFile, xmlFile);
	}
}
