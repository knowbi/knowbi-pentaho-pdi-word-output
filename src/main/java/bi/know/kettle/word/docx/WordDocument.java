package bi.know.kettle.word.docx;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.StringReader;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

public class WordDocument
{
  private static final String DOCUMENT_ENTRY_NAME = "word/document.xml";
  private static final String HEADER_ENTRY_NAME = "word/header\\d+\\.xml";
  private static final String FOOTER_ENTRY_NAME = "word/footer\\d+\\.xml";
  private static final String CHART_ENTRY_NAME = "word/charts/chart\\d+\\.xml";
  private static final String PARAGRAPH_NODE_NAME = "w:p";
  private static final String RULE_NODE_NAME = "w:r";
  private static final String TEXT_NODE_NAME = "w:t";
  private static final String TABLE_NODE_NAME = "w:tbl";
  private static final String TABLE_ROW_NODE_NAME = "w:tr";
  private static final String TABLE_COLUMN_NODE_NAME = "w:tc";
  private static final int INDEX_NOT_FOUND = -1;
  private static final String TEXT_FORMAT = "UTF-8";
  private final HashMap<String, ByteArrayOutputStream> zipEntries;
  private final Document doc;
  private final HashMap<String, Document> headersFooters;
  private final HashMap<String, WordChart> charts;
  private String entryName, end, embed;
  private StringBuilder sb;
  private int start, replLength;
  private ArrayList<Integer> matches;
  
  public WordDocument(InputStream in)
    throws IOException, SAXException, ParserConfigurationException
  {
    this.zipEntries = new HashMap();
    
    ZipInputStream zis = new ZipInputStream(in);
    
    byte[] buffer = new byte[' '];
    try
    {
      ZipEntry entry;
      while ((entry = zis.getNextEntry()) != null)
      {
        String name = entry.getName();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        int len;
        while ((len = zis.read(buffer)) > 0) {
          baos.write(buffer, 0, len);
        }
        this.zipEntries.put(name, baos);
      }
    }
    finally
    {
      zis.close();
    }
    ByteArrayOutputStream baos = (ByteArrayOutputStream)this.zipEntries.get("word/document.xml");
    if (baos == null) {
      throw new IOException("Invalid Document");
    }
    this.doc = getDocFromString(new String(baos.toByteArray(), "UTF-8"));
    
    this.headersFooters = new HashMap();
    this.charts = new HashMap();
    for (Iterator i$ = this.zipEntries.keySet().iterator(); i$.hasNext();)
    {
      entryName = (String)i$.next();
      if ((entryName.matches("word/header\\d+\\.xml")) || (entryName.matches("word/footer\\d+\\.xml")))
      {
        this.headersFooters.put(entryName, getDocFromString(new String(((ByteArrayOutputStream)this.zipEntries.get(entryName)).toByteArray(), "UTF-8")));
      }
      else if (entryName.matches("word/charts/chart\\d+\\.xml"))
      {
        end = entryName.replaceAll(".*\\/(.*?)", "$1");
        for (String relEntryName : this.zipEntries.keySet()) {
          if (relEntryName.equals("word/charts/_rels/" + end + ".rels"))
          {
            String relDoc = new String(((ByteArrayOutputStream)this.zipEntries.get(relEntryName)).toByteArray(), "UTF-8").replace("\n", "").replace("\r", "");
            embed = relDoc.matches(".*?([A-Za-z0-9_\\-]+\\.xlsx).*") ? relDoc.replaceAll(".*?([A-Za-z0-9_\\-]+\\.xlsx).*", "$1") : null;
            for (String embedEntryName : this.zipEntries.keySet()) {
              if (embedEntryName.equals("word/embeddings/" + embed)) {
                this.charts.put(entryName, new WordChart(getDocFromString(new String(((ByteArrayOutputStream)this.zipEntries.get(entryName)).toByteArray(), "UTF-8")), embedEntryName, new XSSFWorkbook(new ByteArrayInputStream(((ByteArrayOutputStream)this.zipEntries.get(embedEntryName)).toByteArray()))));
              }
            }
          }
        }
      }
    }
    String entryName;
    String end;
    String embed;
  }
  
  private synchronized void updateDoc()
    throws TransformerException, IOException
  {
    this.zipEntries.remove("word/document.xml");
    ByteArrayOutputStream baos = new ByteArrayOutputStream();
    baos.write(getStringFromNode(this.doc).getBytes("UTF-8"));
    this.zipEntries.put("word/document.xml", baos);
    for (String entryName : this.headersFooters.keySet())
    {
      this.zipEntries.remove(entryName);
      baos = new ByteArrayOutputStream();
      baos.write(getStringFromNode((Node)this.headersFooters.get(entryName)).getBytes("UTF-8"));
      this.zipEntries.put(entryName, baos);
    }
    for (String entryName : this.charts.keySet())
    {
      this.zipEntries.remove(entryName);
      WordChart c = (WordChart)this.charts.get(entryName);
      baos = new ByteArrayOutputStream();
      baos.write(getStringFromNode(c.getChart()).getBytes("UTF-8"));
      this.zipEntries.put(entryName, baos);
      
      this.zipEntries.remove(c.getWorksheetEntry());
      baos = new ByteArrayOutputStream();
      c.getWorksheet().write(baos);
      this.zipEntries.put(c.getWorksheetEntry(), baos);
    }
  }
  
  public void replaceInText(String search, String replace)
  {
    List<Node> nodeList = new ArrayList();
    
    NodeList nl1 = this.doc.getElementsByTagName("w:p");
    for (int i = 0; i < nl1.getLength(); i++) {
      nodeList.add(nl1.item(i));
    }
    for (String headerFooterName : this.headersFooters.keySet())
    {
      Document headerFooter = (Document)this.headersFooters.get(headerFooterName);
      NodeList nld = headerFooter.getElementsByTagName("w:p");
      for (int i = 0; i < nld.getLength(); i++) {
        nodeList.add(nld.item(i));
      }
    }
    for (Node node : nodeList)
    {
      NodeList nl3 = node.getChildNodes();
      sb = new StringBuilder();
      List<TextNode> textNodeList = new ArrayList();
      start = 0;
      for (int k = 0; k < nl3.getLength(); k++)
      {
        Node node3 = nl3.item(k);
        NodeList nl4 = node3.getChildNodes();
        for (int l = 0; l < nl4.getLength(); l++)
        {
          Node node4 = nl4.item(l);
          if ("w:t".equals(node4.getNodeName()))
          {
            String textPiece = node4.getTextContent();
            sb.append(textPiece);
            textNodeList.add(new TextNode(node4, start, start + textPiece.length()));
            start += textPiece.length();
          }
        }
      }
      if (!textNodeList.isEmpty())
      {
        String text = sb.toString();
        if ((text != null) && (!text.isEmpty()))
        {
          start = 0;
          int end = text.indexOf(search, start);
          if (end != -1)
          {
            matches = new ArrayList<Integer>();
            replLength = search.length();
            while (end != -1)
            {
              matches.add(Integer.valueOf(end));
              start = end + replLength;
              end = text.indexOf(search, start);
            }
            if (!matches.isEmpty()) {
              for (TextNode tn : textNodeList)
              {
                String nodeText = tn.node.getTextContent();
                sb = new StringBuilder();
                start = 0;
                List<Integer> nodeMatches = new ArrayList();
                for (Integer match : matches) {
                  if (((tn.start <= match.intValue()) && (match.intValue() < tn.end)) || ((tn.start > match.intValue()) && (match.intValue() + replLength > tn.start) && (match.intValue() + replLength < tn.end)) || ((tn.start > match.intValue()) && (match.intValue() + replLength >= tn.end))) {
                    nodeMatches.add(match);
                  }
                }
                if (!nodeMatches.isEmpty())
                {
                  for (Integer match : nodeMatches) {
                    if ((tn.start > match.intValue()) && (match.intValue() + replLength >= tn.end))
                    {
                      start = nodeText.length();
                    }
                    else if ((tn.start > match.intValue()) && (match.intValue() + replLength > tn.start) && (match.intValue() + replLength < tn.end))
                    {
                      start = match.intValue() + replLength - tn.start;
                    }
                    else if ((tn.start <= match.intValue()) && (match.intValue() < tn.end))
                    {
                      sb.append(nodeText.substring(start, match.intValue() - tn.start)).append(replace);
                      
                      start = match.intValue() + replLength - tn.start;
                      if (start > nodeText.length()) {
                        start = nodeText.length();
                      }
                    }
                  }
                  sb.append(nodeText.substring(start));
                  tn.node.setTextContent(sb.toString());
                }
              }
            }
          }
        }
      }
    }
    StringBuilder sb;
    int start;
    List<Integer> matches;
    int replLength;
    for (String chartString : this.charts.keySet())
    {
      WordChart chart = (WordChart)this.charts.get(chartString);
      chart.replaceInTitle(search, replace);
      chart.replaceInSerieNames(search, replace);
    }
  }
  
  public int getNumberOfTables()
  {
    NodeList nl = this.doc.getElementsByTagName("w:tbl");
    return nl.getLength();
  }
  
  public String getStringTableContent(int table)
  {
    StringBuilder sb = new StringBuilder();
    NodeList nl1 = this.doc.getElementsByTagName("w:tbl");
    int tableNr = 0;
    for (int i = 0; i < nl1.getLength(); i++)
    {
      tableNr += 1;
      if (tableNr == table)
      {
        Node node1 = nl1.item(i);
        NodeList nl2 = this.doc.getElementsByTagName("w:p");
        for (int j = 0; j < nl2.getLength(); j++)
        {
          Node node2 = nl2.item(j);
          Node pNode = node2;
          int cnt = 0;
          while ((cnt < 3) && (pNode != null))
          {
            cnt++;
            pNode = pNode.getParentNode();
          }
          if ((pNode != null) && (pNode == node1))
          {
            NodeList nl3 = node2.getChildNodes();
            for (int k = 0; k < nl3.getLength(); k++)
            {
              Node node3 = nl3.item(k);
              NodeList nl4 = node3.getChildNodes();
              for (int l = 0; l < nl4.getLength(); l++)
              {
                Node node4 = nl4.item(l);
                sb.append(node4.getTextContent());
              }
            }
            sb.append("|");
          }
        }
      }
    }
    if (sb.length() > 0) {
      sb.deleteCharAt(sb.length() - 1);
    }
    return sb.toString();
  }
  
  public void insertInTable(int table, int row, int startcol, String[] fields, boolean appendRow)
  {
    NodeList nl1 = this.doc.getElementsByTagName("w:tbl");
    int tableNr = 0;
    for (int i = 0; i < nl1.getLength(); i++)
    {
      tableNr += 1;
      if (tableNr == table)
      {
        Node node1 = nl1.item(i);
        NodeList nl2 = node1.getChildNodes();
        int rowNr = 0;
        Node appendNode = null;
        for (int j = 0; j < nl2.getLength(); j++)
        {
          Node node2 = nl2.item(j);
          if ("w:tr".equals(node2.getNodeName()))
          {
            rowNr += 1;
            appendNode = node2;
            if (row == rowNr) {
              insertInTableRow(node2, startcol, fields);
            }
          }
        }
        if ((appendRow) && (appendNode != null) && (rowNr < row))
        {
          Node node2 = appendNode.cloneNode(true);
          node1.appendChild(node2);
          insertInTableRow(node2, startcol, fields);
        }
      }
    }
  }
  
  private void insertInTableRow(Node rowNode, int startcol, String[] fields)
  {
    NodeList nl3 = rowNode.getChildNodes();
    int colNr = 0;
    for (int k = 0; k < nl3.getLength(); k++)
    {
      Node node3 = nl3.item(k);
      if ("w:tc".equals(node3.getNodeName()))
      {
        colNr += 1;
        if (colNr >= startcol)
        {
          NodeList nl4 = node3.getChildNodes();
          boolean firstPar = true;
          for (int l = 0; l < nl4.getLength(); l++)
          {
            Node node4 = nl4.item(l);
            if ("w:p".equals(node4.getNodeName()))
            {
              if (!firstPar) {
                node3.removeChild(node4);
              }
              firstPar = false;
              NodeList nl5 = node4.getChildNodes();
              boolean firstRule = true;
              int ruleCnt = 0;
              for (int m = 0; m < nl5.getLength(); m++)
              {
                Node node5 = nl5.item(m);
                if ("w:r".equals(node5.getNodeName()))
                {
                  ruleCnt++;
                  if (!firstRule) {
                    node4.removeChild(node5);
                  }
                  firstRule = false;
                  NodeList nl6 = node5.getChildNodes();
                  boolean firstText = true;
                  for (int n = 0; n < nl6.getLength(); n++)
                  {
                    Node node6 = nl6.item(n);
                    if ("w:t".equals(node6.getNodeName()))
                    {
                      if (!firstText) {
                        node5.removeChild(node6);
                      }
                      firstText = false;
                      if ((colNr - startcol < fields.length) && (colNr - startcol >= 0))
                      {
                        String field = fields[(colNr - startcol)];
                        node6.setTextContent(field);
                      }
                    }
                  }
                }
              }
              if ((ruleCnt == 0) && (colNr - startcol < fields.length) && (colNr - startcol >= 0))
              {
                Node node5 = this.doc.createElement("w:r");
                node4.appendChild(node5);
                Node node6 = this.doc.createElement("w:t");
                node5.appendChild(node6);
                node6.setTextContent(fields[(colNr - startcol)]);
              }
            }
          }
        }
      }
    }
  }
  
  public List<String> getChartNames()
  {
    List<String> sl = new ArrayList();
    for (String s : this.charts.keySet()) {
      sl.add(s.replaceAll(".*\\/(.*?)\\.xml", "$1"));
    }
    return sl;
  }
  
  public WordChart getChart(String name)
  {
    for (String s : this.charts.keySet()) {
      if (s.replaceAll(".*\\/(.*?)\\.xml", "$1").matches(name)) {
        return (WordChart)this.charts.get(s);
      }
    }
    return null;
  }
  
  private class TextNode
  {
    private final Node node;
    private final int start;
    private final int end;
    
    private TextNode(Node node, int start, int end)
    {
      this.node = node;
      this.start = start;
      this.end = end;
    }
  }
  
  public void write(OutputStream out)
    throws TransformerException, IOException
  {
    updateDoc();
    
    ZipOutputStream zos = null;
    try
    {
      zos = new ZipOutputStream(out);
      for (String name : this.zipEntries.keySet())
      {
        zos.putNextEntry(new ZipEntry(name));
        zos.write(((ByteArrayOutputStream)this.zipEntries.get(name)).toByteArray());
      }
    }
    finally
    {
      zos.close();
    }
  }
  
  private static Document getDocFromString(String s)
    throws ParserConfigurationException, SAXException, IOException
  {
    DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
    dbf.setNamespaceAware(true);
    DocumentBuilder db = dbf.newDocumentBuilder();
    return db.parse(new InputSource(new StringReader(s)));
  }
  
  public static String getStringFromNode(Node n)
    throws TransformerException
  {
    TransformerFactory transformerFactory = TransformerFactory.newInstance();
    Transformer transformer = transformerFactory.newTransformer();
    DOMSource source = new DOMSource(n);
    StreamResult result = new StreamResult(new StringWriter());
    transformer.transform(source, result);
    return result.getWriter().toString();
  }
}


/* Location:              /home/bart/bart.maertens@know.bi/Projecten/Drivolution/word-plugin/pentaho_wordwriter/wordwriter.jar!/be/drivolution/docx/WordDocument.class
 * Java compiler version: 6 (50.0)
 * JD-Core Version:       0.7.1
 */