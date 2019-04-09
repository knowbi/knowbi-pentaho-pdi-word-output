package bi.know.kettle.word.docx;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class WordChart
{
  private static final String PARAGRAPH_NODE_NAME = "a:p";
  private static final String RULE_NODE_NAME = "a:r";
  private static final String TEXT_NODE_NAME = "a:t";
  private static final String CHART_SERIES_NODE_NAME = "c:ser";
  private static final String NAME_CHART_SERIES_NODE_NAME = "c:tx";
  private static final int INDEX_NOT_FOUND = -1;
  private final Document chart;
  private final String worksheetEntry;
  private final XSSFWorkbook worksheet;
  private final Node title;
  private final Node plotarea;
  
  public WordChart(Document chart, String worksheetEntry, XSSFWorkbook worksheet)
  {
    this.chart = chart;
    this.worksheetEntry = worksheetEntry;
    this.worksheet = worksheet;
    
    NodeList nl = chart.getElementsByTagName("c:title");
    Node n = null;
    for (int i = 0; i < nl.getLength(); i++) {
      n = nl.item(i);
    }
    this.title = n;
    
    n = null;
    nl = chart.getElementsByTagName("c:plotArea");
    for (int i = 0; i < nl.getLength(); i++) {
      n = nl.item(i);
    }
    this.plotarea = n;
  }
  
  public String getTitle()
  {
    StringBuilder sb = new StringBuilder();
    NodeList nl1 = this.title.getChildNodes();
    for (int i = 0; i < nl1.getLength(); i++)
    {
      Node node1 = nl1.item(i);
      NodeList nl2 = node1.getChildNodes();
      for (int j = 0; j < nl2.getLength(); j++)
      {
        Node node2 = nl2.item(j);
        NodeList nl3 = node2.getChildNodes();
        for (int k = 0; k < nl3.getLength(); k++)
        {
          Node node3 = nl3.item(k);
          if (node3.getNodeName().equals("a:p"))
          {
            NodeList nl4 = node3.getChildNodes();
            for (int l = 0; l < nl4.getLength(); l++)
            {
              Node node4 = nl4.item(l);
              if (node4.getNodeName().equals("a:r"))
              {
                NodeList nl5 = node4.getChildNodes();
                for (int m = 0; m < nl5.getLength(); m++)
                {
                  Node node5 = nl5.item(m);
                  if (node5.getNodeName().equals("a:t")) {
                    sb.append(node5.getTextContent()).append(" ");
                  }
                }
              }
            }
          }
        }
      }
    }
    int index = sb.lastIndexOf(" ");
    if (index != -1) {
      sb.deleteCharAt(index);
    }
    return sb.toString();
  }
  
  public void replaceInTitle(String search, String replace)
  {
    NodeList nl1 = this.title.getChildNodes();
    for (int i = 0; i < nl1.getLength(); i++)
    {
      Node node1 = nl1.item(i);
      NodeList nl2 = node1.getChildNodes();
      for (int j = 0; j < nl2.getLength(); j++)
      {
        Node node2 = nl2.item(j);
        NodeList nl3 = node2.getChildNodes();
        StringBuilder sb;
        int start;
        List<Integer> matches;
        int replLength;
        for (int k = 0; k < nl3.getLength(); k++)
        {
          Node node3 = nl3.item(k);
          if ("a:p".equals(node3.getNodeName()))
          {
            NodeList nl4 = node3.getChildNodes();
            sb = new StringBuilder();
            List<TextNode> textNodeList = new ArrayList();
            start = 0;
            for (int m = 0; m < nl4.getLength(); m++)
            {
              Node node4 = nl4.item(m);
              if (node4.getNodeName().equals("a:r"))
              {
                NodeList nl5 = node4.getChildNodes();
                for (int n = 0; n < nl5.getLength(); n++)
                {
                  Node node5 = nl5.item(n);
                  if ("a:t".equals(node5.getNodeName()))
                  {
                    String textPiece = node5.getTextContent();
                    sb.append(textPiece);
                    textNodeList.add(new TextNode(node5, start, start + textPiece.length()));
                    start += textPiece.length();
                  }
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
                  matches = new ArrayList();
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
        }
      }
    }
  }
  
  public void replaceInSerieNames(String search, String replace)
  {
    NodeList nl1 = this.plotarea.getChildNodes();
    for (int i = 0; i < nl1.getLength(); i++)
    {
      Node node1 = nl1.item(i);
      NodeList nl2 = node1.getChildNodes();
      for (int j = 0; j < nl2.getLength(); j++)
      {
        Node node2 = nl2.item(j);
        if ("c:ser".equals(node2.getNodeName()))
        {
          NodeList nl3 = node2.getChildNodes();
          for (int k = 0; k < nl3.getLength(); k++)
          {
            Node node3 = nl3.item(k);
            if ("c:tx".equals(node3.getNodeName()))
            {
              NodeList nl4 = node3.getChildNodes();
              for (int l = 0; l < nl4.getLength(); l++)
              {
                Node node4 = nl4.item(l);
                if ("c:strRef".equals(node4.getNodeName()))
                {
                  NodeList nl5 = node4.getChildNodes();
                  String s = null;
                  for (int m = 0; m < nl5.getLength(); m++)
                  {
                    Node node5 = nl5.item(m);
                    if ("c:f".equals(node5.getNodeName())) {
                      s = replaceInWorksheet(node5.getTextContent(), search, replace);
                    }
                  }
                  if (s != null) {
                    for (int m = 0; m < nl5.getLength(); m++)
                    {
                      Node node5 = nl5.item(m);
                      if ("c:strCache".equals(node5.getNodeName()))
                      {
                        NodeList nl6 = node5.getChildNodes();
                        for (int n = 0; n < nl6.getLength(); n++)
                        {
                          Node node6 = nl6.item(n);
                          if ("c:pt".equals(node6.getNodeName()))
                          {
                            NodeList nl7 = node6.getChildNodes();
                            for (int o = 0; o < nl7.getLength(); o++)
                            {
                              Node node7 = nl7.item(o);
                              if ("c:v".equals(node7.getNodeName())) {
                                node7.setTextContent(s);
                              }
                            }
                          }
                        }
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }
  
  private String replaceInWorksheet(String ref, String search, String replace)
  {
    String sheetName = ref.matches("\\w+!.*") ? ref.replaceFirst("(\\w+)!.*", "$1") : null;
    String colStr = ref.matches(".*!.*?[A-Z]+.*") ? ref.replaceFirst(".*!.*?([A-Z]+).*", "$1") : null;
    String rowStr = ref.matches(".*!.*?\\d+") ? ref.replaceFirst(".*!.*?(\\d+)", "$1") : null;
    if ((sheetName == null) || (colStr == null) || (rowStr == null)) {
      return null;
    }
    int col = charToInt(colStr);
    int row = Integer.parseInt(rowStr);
    if (this.worksheet == null) {
      return null;
    }
    XSSFSheet sheet = this.worksheet.getSheet(sheetName);
    if (sheet == null) {
      return null;
    }
    XSSFCell cell = sheet.getRow(row - 1).getCell(col - 1);
    if (cell == null) {
      return null;
    }
    String s = cell.getRichStringCellValue().getString();
    s = s.replace(search, replace);
    cell.getRichStringCellValue().setString(s);
    return s;
  }
  
  private int charToInt(String s)
  {
    char[] c = s.toUpperCase().toCharArray();
    char[] d = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".toCharArray();
    int res = 0;
    for (int i = 0; i < c.length; i++) {
      for (int j = 0; j < d.length; j++) {
        if (c[(c.length - 1 - i)] == d[j])
        {
          res += (int)Math.pow(26.0D, i) * (j + 1);
          break;
        }
      }
    }
    return res;
  }
  
  public void insertRow(int row, Number[] values, boolean updateLabel, String label, boolean appendRows)
  {
    NodeList nl1 = this.plotarea.getChildNodes();
    int serCnt = 0;
    for (int i = 0; i < nl1.getLength(); i++)
    {
      Node node1 = nl1.item(i);
      NodeList nl2 = node1.getChildNodes();
      for (int j = 0; j < nl2.getLength(); j++)
      {
        Node node2 = nl2.item(j);
        if ("c:ser".equals(node2.getNodeName()))
        {
          NodeList nl3 = node2.getChildNodes();
          for (int k = 0; k < nl3.getLength(); k++)
          {
            Node node3 = nl3.item(k);
            if ((updateLabel) && ("c:cat".equals(node3.getNodeName())))
            {
              NodeList nl4 = node3.getChildNodes();
              for (int l = 0; l < nl4.getLength(); l++)
              {
                Node node4 = nl4.item(l);
                if ("c:strRef".equals(node4.getNodeName()))
                {
                  NodeList nl5 = node4.getChildNodes();
                  for (int m = 0; m < nl5.getLength(); m++)
                  {
                    Node node5 = nl5.item(m);
                    if ("c:f".equals(node5.getNodeName()))
                    {
                      String s = replaceStringInWorksheet(node5.getTextContent(), row, label, appendRows);
                      if (s != null) {
                        node5.setTextContent(s);
                      }
                    }
                  }
                  for (int m = 0; m < nl5.getLength(); m++)
                  {
                    Node node5 = nl5.item(m);
                    if ("c:strCache".equals(node5.getNodeName()))
                    {
                      NodeList nl6 = node5.getChildNodes();
                      int cnt = 0;
                      Node ptNode = null;
                      for (int n = 0; n < nl6.getLength(); n++)
                      {
                        Node node6 = nl6.item(n);
                        if ("c:pt".equals(node6.getNodeName()))
                        {
                          ptNode = node6;
                          cnt++;
                          if (row == cnt)
                          {
                            NodeList nl7 = node6.getChildNodes();
                            for (int o = 0; o < nl7.getLength(); o++)
                            {
                              Node node7 = nl7.item(o);
                              if ("c:v".equals(node7.getNodeName())) {
                                node7.setTextContent(label);
                              }
                            }
                          }
                        }
                      }
                      while ((appendRows) && (cnt < row) && (ptNode != null))
                      {
                        Node node6 = ptNode.cloneNode(true);
                        node5.appendChild(node6);
                        ((Element)node6).setAttribute("idx", Integer.toString(cnt));
                        cnt++;
                        if (row == cnt)
                        {
                          NodeList nl7 = node6.getChildNodes();
                          for (int n = 0; n < nl7.getLength(); n++)
                          {
                            Node node7 = nl7.item(n);
                            if ("c:v".equals(node7.getNodeName())) {
                              node7.setTextContent(label);
                            }
                          }
                        }
                      }
                    }
                  }
                }
              }
            }
            if ("c:val".equals(node3.getNodeName()))
            {
              NodeList nl4 = node3.getChildNodes();
              for (int l = 0; l < nl4.getLength(); l++)
              {
                Node node4 = nl4.item(l);
                if ("c:numRef".equals(node4.getNodeName()))
                {
                  NodeList nl5 = node4.getChildNodes();
                  if (values.length > serCnt)
                  {
                    Number num = values[serCnt];
                    String numStr = num != null ? num.toString() : "";
                    for (int m = 0; m < nl5.getLength(); m++)
                    {
                      Node node5 = nl5.item(m);
                      if ("c:f".equals(node5.getNodeName()))
                      {
                        String s = replaceNumberInWorksheet(node5.getTextContent(), row, num, (appendRows) && (updateLabel));
                        if (s != null) {
                          node5.setTextContent(s);
                        }
                      }
                    }
                    for (int m = 0; m < nl5.getLength(); m++)
                    {
                      Node node5 = nl5.item(m);
                      if ("c:numCache".equals(node5.getNodeName()))
                      {
                        NodeList nl6 = node5.getChildNodes();
                        int cnt = 0;
                        Node ptNode = null;
                        for (int n = 0; n < nl6.getLength(); n++)
                        {
                          Node node6 = nl6.item(n);
                          if ("c:pt".equals(node6.getNodeName()))
                          {
                            cnt++;
                            ptNode = node6;
                            if (row == cnt)
                            {
                              NodeList nl7 = node6.getChildNodes();
                              for (int o = 0; o < nl7.getLength(); o++)
                              {
                                Node node7 = nl7.item(o);
                                if ("c:v".equals(node7.getNodeName())) {
                                  node7.setTextContent(numStr);
                                }
                              }
                            }
                          }
                        }
                        while ((appendRows) && (updateLabel) && (cnt < row) && (ptNode != null))
                        {
                          Node node6 = ptNode.cloneNode(true);
                          node5.appendChild(node6);
                          ((Element)node6).setAttribute("idx", Integer.toString(cnt));
                          cnt++;
                          if (row == cnt)
                          {
                            NodeList nl7 = node6.getChildNodes();
                            for (int n = 0; n < nl7.getLength(); n++)
                            {
                              Node node7 = nl7.item(n);
                              if ("c:v".equals(node7.getNodeName())) {
                                node7.setTextContent(numStr);
                              }
                            }
                          }
                        }
                      }
                    }
                  }
                }
              }
            }
          }
          serCnt++;
        }
      }
    }
  }
  
  private String replaceStringInWorksheet(String ref, int row, String label, boolean appendRow)
  {
    String sheetName = ref.matches("\\w+!.*") ? ref.replaceFirst("(\\w+)!.*", "$1") : null;
    String colStr = ref.matches(".*!.*?[A-Z]+.*") ? ref.replaceFirst(".*!.*?([A-Z]+).*", "$1") : null;
    String rowStr1 = ref.matches(".*!.*?\\d+.*?\\d+") ? ref.replaceFirst(".*!.*?(\\d+).*?(\\d+)", "$1") : null;
    if (rowStr1 == null) {
      rowStr1 = ref.matches(".*!.*?\\d+") ? ref.replaceFirst(".*!.*?(\\d+)", "$1") : null;
    }
    String rowStr2 = ref.matches(".*!.*?\\d+.*?\\d+") ? ref.replaceFirst(".*!.*?(\\d+).*?(\\d+)", "$2") : null;
    if (rowStr2 == null) {
      rowStr2 = rowStr1;
    }
    if ((sheetName == null) || (colStr == null) || (rowStr1 == null) || (rowStr2 == null)) {
      return null;
    }
    int col = charToInt(colStr);
    int row1 = Integer.parseInt(rowStr1);
    int row2 = Integer.parseInt(rowStr2);
    
    int rrow = row1 + row - 1;
    if (this.worksheet == null) {
      return null;
    }
    XSSFSheet sheet = this.worksheet.getSheet(sheetName);
    if (sheet == null) {
      return null;
    }
    XSSFRow xrow = sheet.getRow(rrow - 1);
    if ((xrow == null) && (appendRow)) {
      xrow = sheet.createRow(rrow - 1);
    } else if (xrow == null) {
      return null;
    }
    XSSFCell cell = xrow.getCell(col - 1);
    if (cell == null) {
      cell = xrow.createCell(col - 1);
    }
    cell.setCellValue(label);
    if (row2 >= rrow) {
      return ref;
    }
    if (appendRow) {
      return sheetName + "!$" + colStr + "$" + rowStr1 + ":$" + colStr + "$" + rrow;
    }
    return null;
  }
  
  private String replaceNumberInWorksheet(String ref, int row, Number value, boolean appendRow)
  {
    String sheetName = ref.matches("\\w+!.*") ? ref.replaceFirst("(\\w+)!.*", "$1") : null;
    String colStr = ref.matches(".*!.*?[A-Z]+.*") ? ref.replaceFirst(".*!.*?([A-Z]+).*", "$1") : null;
    String rowStr1 = ref.matches(".*!.*?\\d+.*?\\d+") ? ref.replaceFirst(".*!.*?(\\d+).*?(\\d+)", "$1") : null;
    if (rowStr1 == null) {
      rowStr1 = ref.matches(".*!.*?\\d+") ? ref.replaceFirst(".*!.*?(\\d+)", "$1") : null;
    }
    String rowStr2 = ref.matches(".*!.*?\\d+.*?\\d+") ? ref.replaceFirst(".*!.*?(\\d+).*?(\\d+)", "$2") : null;
    if (rowStr2 == null) {
      rowStr2 = rowStr1;
    }
    if ((sheetName == null) || (colStr == null) || (rowStr1 == null) || (rowStr2 == null)) {
      return null;
    }
    int col = charToInt(colStr);
    int row1 = Integer.parseInt(rowStr1);
    int row2 = Integer.parseInt(rowStr2);
    
    int rrow = row1 + row - 1;
    if (this.worksheet == null) {
      return null;
    }
    XSSFSheet sheet = this.worksheet.getSheet(sheetName);
    if (sheet == null) {
      return null;
    }
    XSSFRow xrow = sheet.getRow(rrow - 1);
    if ((xrow == null) && (appendRow)) {
      xrow = sheet.createRow(rrow - 1);
    } else if (xrow == null) {
      return null;
    }
    XSSFCell cell = xrow.getCell(col - 1);
    if (cell == null) {
      cell = xrow.createCell(col - 1);
    }
    if (value == null) {
      xrow.removeCell(cell);
    } else {
      cell.setCellValue(value.doubleValue());
    }
    if (row2 >= rrow) {
      return ref;
    }
    if (appendRow) {
      return sheetName + "!$" + colStr + "$" + rowStr1 + ":$" + colStr + "$" + rrow;
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
  
  public Document getChart()
  {
    return this.chart;
  }
  
  public String getWorksheetEntry()
  {
    return this.worksheetEntry;
  }
  
  public XSSFWorkbook getWorksheet()
  {
    return this.worksheet;
  }
}


/* Location:              /home/bart/bart.maertens@know.bi/Projecten/Drivolution/word-plugin/pentaho_wordwriter/wordwriter.jar!/be/drivolution/docx/WordChart.class
 * Java compiler version: 6 (50.0)
 * JD-Core Version:       0.7.1
 */