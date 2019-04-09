package bi.know.kettle.word.output;

import java.util.List;
import java.util.Map;
import org.eclipse.swt.widgets.Shell;
import org.pentaho.di.core.CheckResult;
import org.pentaho.di.core.CheckResultInterface;
import org.pentaho.di.core.Counter;
import org.pentaho.di.core.annotations.Step;
import org.pentaho.di.core.database.DatabaseMeta;
import org.pentaho.di.core.exception.KettleException;
import org.pentaho.di.core.exception.KettleXMLException;
import org.pentaho.di.core.row.RowMetaInterface;
import org.pentaho.di.core.variables.VariableSpace;
import org.pentaho.di.core.xml.XMLHandler;
import org.pentaho.di.i18n.BaseMessages;
import org.pentaho.di.repository.ObjectId;
import org.pentaho.di.repository.Repository;
import org.pentaho.di.trans.Trans;
import org.pentaho.di.trans.TransMeta;
import org.pentaho.di.trans.step.BaseStepMeta;
import org.pentaho.di.trans.step.StepDataInterface;
import org.pentaho.di.trans.step.StepDialogInterface;
import org.pentaho.di.trans.step.StepInterface;
import org.pentaho.di.trans.step.StepMeta;
import org.pentaho.di.trans.step.StepMetaInterface;
import org.w3c.dom.Node;


@Step(id = "WordOutput",
        image = "NEO4J.svg",
        i18nPackageName="bi.know.kettle.word.output",
        name="WordOutput.Step.Name",
        description = "WordOutput.Step.Description",
        categoryDescription="WordOutput.Step.Category",
        isSeparateClassLoaderNeeded=true
        )

public class WordWriterMeta
  extends BaseStepMeta
  implements StepMetaInterface
{
  private static Class<?> PKG = WordWriterMeta.class;
  public static final String FUNCTION_REPLACE_IN_TEXT = "replace_in_text";
  public static final String FUNCTION_INSERT_IN_TABLE = "insert_in_table";
  public static final String FUNCTION_INSERT_IN_CHART = "insert_in_chart";
  private String filename;
  private String function;
  private String search;
  private String replaceBy;
  private int table;
  private int startRow;
  private int startColumn;
  private boolean appendRows;
  private String[] tableColumns;
  private String chart;
  private boolean updateLabel;
  private String label;
  private boolean cAppendRows;
  private String[] chartColumns;
  
  public StepInterface getStep(StepMeta stepMeta, StepDataInterface stepDataInterface, int cnr, TransMeta tr, Trans trans)
  {
    return new WordWriter(stepMeta, stepDataInterface, cnr, tr, trans);
  }
  
  public StepDialogInterface getDialog(Shell shell, StepMetaInterface meta, TransMeta transMeta, String name)
  {
    return new WordWriterDialog(shell, meta, transMeta, name);
  }
  
  public void setDefault()
  {
    this.filename = "";
    this.function = "replace_in_text";
    
    this.search = "";
    this.replaceBy = "";
    
    this.table = 0;
    this.startRow = 1;
    this.startColumn = 1;
    this.appendRows = false;
    this.tableColumns = new String[0];
    
    this.chart = "";
    this.updateLabel = false;
    this.label = "";
    this.cAppendRows = false;
    this.chartColumns = new String[0];
  }
  
  public StepDataInterface getStepData()
  {
    return new WordWriterData();
  }
  
  public void saveRep(Repository rep, ObjectId id_transformation, ObjectId id_step)
    throws KettleException
  {
    try
    {
      rep.saveStepAttribute(id_transformation, id_step, "filename", this.filename);
      rep.saveStepAttribute(id_transformation, id_step, "function", this.function);
      rep.saveStepAttribute(id_transformation, id_step, "search", this.search);
      rep.saveStepAttribute(id_transformation, id_step, "replace_by", this.replaceBy);
      rep.saveStepAttribute(id_transformation, id_step, "table", this.table);
      rep.saveStepAttribute(id_transformation, id_step, "start_row", this.startRow);
      rep.saveStepAttribute(id_transformation, id_step, "start_column", this.startColumn);
      rep.saveStepAttribute(id_transformation, id_step, "append_rows", this.appendRows);
      for (int i = 0; i < this.tableColumns.length; i++) {
        rep.saveStepAttribute(id_transformation, id_step, i, "table_column", this.tableColumns[i]);
      }
      rep.saveStepAttribute(id_transformation, id_step, "chart", this.chart);
      rep.saveStepAttribute(id_transformation, id_step, "update_label", this.updateLabel);
      rep.saveStepAttribute(id_transformation, id_step, "label", this.label);
      rep.saveStepAttribute(id_transformation, id_step, "c_append_rows", this.cAppendRows);
      for (int i = 0; i < this.chartColumns.length; i++) {
        rep.saveStepAttribute(id_transformation, id_step, i, "chart_column", this.chartColumns[i]);
      }
    }
    catch (Exception e)
    {
      throw new KettleException("Unable to save step into repository: " + id_step, e);
    }
  }
  
  public void check(List<CheckResultInterface> remarks, TransMeta transmeta, StepMeta stepMeta, RowMetaInterface prev, String[] input, String[] output, RowMetaInterface info)
  {
    if (input.length > 0)
    {
      CheckResult cr = new CheckResult(1, BaseMessages.getString(PKG, "Demo.CheckResult.ReceivingRows.OK", new String[0]), stepMeta);
      remarks.add(cr);
    }
    else
    {
      CheckResult cr = new CheckResult(4, BaseMessages.getString(PKG, "Demo.CheckResult.ReceivingRows.ERROR", new String[0]), stepMeta);
      remarks.add(cr);
    }
  }
  
  public void readRep(Repository rep, ObjectId id_step, List<DatabaseMeta> databases, Map<String, Counter> counters)
    throws KettleException
  {
    try
    {
      this.filename = rep.getStepAttributeString(id_step, "filename");
      this.function = rep.getStepAttributeString(id_step, "function");
      this.search = rep.getStepAttributeString(id_step, "search");
      this.replaceBy = rep.getStepAttributeString(id_step, "replace_by");
      this.table = ((int)rep.getStepAttributeInteger(id_step, "table"));
      this.startRow = ((int)rep.getStepAttributeInteger(id_step, "start_row"));
      this.startColumn = ((int)rep.getStepAttributeInteger(id_step, "start_column"));
      this.appendRows = rep.getStepAttributeBoolean(id_step, "append_rows");
      
      int nrfields = rep.countNrStepAttributes(id_step, "table_column");
      this.tableColumns = new String[nrfields];
      for (int i = 0; i < nrfields; i++) {
        this.tableColumns[i] = rep.getStepAttributeString(id_step, i, "table_column");
      }
      this.chart = rep.getStepAttributeString(id_step, "chart");
      this.updateLabel = rep.getStepAttributeBoolean(id_step, "update_label");
      this.label = rep.getStepAttributeString(id_step, "label");
      this.cAppendRows = rep.getStepAttributeBoolean(id_step, "c_append_rows");
      
      nrfields = rep.countNrStepAttributes(id_step, "chart_column");
      this.chartColumns = new String[nrfields];
      for (int i = 0; i < nrfields; i++) {
        this.chartColumns[i] = rep.getStepAttributeString(id_step, i, "table_column");
      }
    }
    catch (Exception e)
    {
      throw new KettleException("Unable to load step from repository", e);
    }
  }
  
  public String getXML()
    throws KettleException
  {
    String xml = XMLHandler.addTagValue("filename", this.filename);
    xml = xml + XMLHandler.addTagValue("function", this.function);
    xml = xml + XMLHandler.addTagValue("search", this.search);
    xml = xml + XMLHandler.addTagValue("replace_by", this.replaceBy);
    xml = xml + XMLHandler.addTagValue("table", this.table);
    xml = xml + XMLHandler.addTagValue("start_row", this.startRow);
    xml = xml + XMLHandler.addTagValue("start_column", this.startColumn);
    xml = xml + XMLHandler.addTagValue("append_rows", this.appendRows);
    xml = xml + "<table_columns>";
    for (String tableColumn : this.tableColumns) {
      xml = xml + "  " + XMLHandler.addTagValue("table_column", tableColumn);
    }
    xml = xml + "</table_columns>";
    xml = xml + XMLHandler.addTagValue("chart", this.chart);
    xml = xml + XMLHandler.addTagValue("update_label", this.updateLabel);
    xml = xml + XMLHandler.addTagValue("label", this.label);
    xml = xml + XMLHandler.addTagValue("c_append_rows", this.cAppendRows);
    xml = xml + "<chart_columns>";
    for (String chartColumn : this.chartColumns) {
      xml = xml + "  " + XMLHandler.addTagValue("chart_column", chartColumn);
    }
    xml = xml + "</chart_columns>";
    return xml;
  }
  
  public void loadXML(Node stepnode, List<DatabaseMeta> databases, Map<String, Counter> counters)
    throws KettleXMLException
  {
    try
    {
      setFilename(XMLHandler.getNodeValue(XMLHandler.getSubNode(stepnode, "filename")));
      setFunction(XMLHandler.getNodeValue(XMLHandler.getSubNode(stepnode, "function")));
      setSearch(XMLHandler.getNodeValue(XMLHandler.getSubNode(stepnode, "search")));
      setReplaceBy(XMLHandler.getNodeValue(XMLHandler.getSubNode(stepnode, "replace_by")));
      
      String tableValue = XMLHandler.getNodeValue(XMLHandler.getSubNode(stepnode, "table"));
      if (tableValue != null) {
        setTable(Integer.parseInt(tableValue));
      } else {
        setTable(0);
      }
      String startRow = XMLHandler.getNodeValue(XMLHandler.getSubNode(stepnode, "start_row"));
      if (startRow != null) {
        setStartRow(Integer.parseInt(startRow));
      } else {
        setStartRow(1);
      }
      String startColumn = XMLHandler.getNodeValue(XMLHandler.getSubNode(stepnode, "start_column"));
      if (startColumn != null) {
        setStartColumn(Integer.parseInt(startColumn));
      } else {
        setStartColumn(1);
      }
      setAppendRows("Y".equalsIgnoreCase(XMLHandler.getNodeValue(XMLHandler.getSubNode(stepnode, "append_rows"))));
      
      Node fields = XMLHandler.getSubNode(stepnode, "table_columns");
      int nrfields = XMLHandler.countNodes(fields, "table_column");
      this.tableColumns = new String[nrfields];
      for (int i = 0; i < nrfields; i++)
      {
        this.tableColumns[i] = XMLHandler.getNodeValue(XMLHandler.getSubNodeByNr(fields, "table_column", i));
        if (this.tableColumns[i] == null) {
          this.tableColumns[i] = "";
        }
      }
      setChart(XMLHandler.getNodeValue(XMLHandler.getSubNode(stepnode, "chart")));
      setUpdateLabel("Y".equals(XMLHandler.getNodeValue(XMLHandler.getSubNode(stepnode, "update_label"))));
      setLabel(XMLHandler.getNodeValue(XMLHandler.getSubNode(stepnode, "label")));
      setCAppendRows("Y".equals(XMLHandler.getNodeValue(XMLHandler.getSubNode(stepnode, "c_append_rows"))));
      
      fields = XMLHandler.getSubNode(stepnode, "chart_columns");
      nrfields = XMLHandler.countNodes(fields, "chart_column");
      this.chartColumns = new String[nrfields];
      for (int i = 0; i < nrfields; i++)
      {
        this.chartColumns[i] = XMLHandler.getNodeValue(XMLHandler.getSubNodeByNr(fields, "chart_column", i));
        if (this.chartColumns[i] == null) {
          this.chartColumns[i] = "";
        }
      }
      if (getFilename() == null) {
        setFilename("");
      }
      if (getFunction() == null) {
        setFunction("replace_in_text");
      }
      if (getSearch() == null) {
        setSearch("");
      }
      if (getReplaceBy() == null) {
        setReplaceBy("");
      }
    }
    catch (Exception e)
    {
      throw new KettleXMLException("Word Writer plugin unable to read step info from XML node", e);
    }
  }
  
  public void setFilename(String filename)
  {
    this.filename = filename;
  }
  
  public String getFilename()
  {
    return this.filename;
  }
  
  public String buildFilename(VariableSpace space)
  {
    return space.environmentSubstitute(this.filename);
  }
  
  public void setFunction(String function)
  {
    this.function = function;
  }
  
  public String getFunction()
  {
    return this.function;
  }
  
  public void setSearch(String search)
  {
    this.search = search;
  }
  
  public String getSearch()
  {
    return this.search;
  }
  
  public void setReplaceBy(String replaceBy)
  {
    this.replaceBy = replaceBy;
  }
  
  public String getReplaceBy()
  {
    return this.replaceBy;
  }
  
  public void setTable(int table)
  {
    this.table = table;
  }
  
  public int getTable()
  {
    return this.table;
  }
  
  public void setStartRow(int startRow)
  {
    this.startRow = startRow;
  }
  
  public int getStartRow()
  {
    return this.startRow;
  }
  
  public void setStartColumn(int startColumn)
  {
    this.startColumn = startColumn;
  }
  
  public int getStartColumn()
  {
    return this.startColumn;
  }
  
  public void setAppendRows(boolean appendRows)
  {
    this.appendRows = appendRows;
  }
  
  public boolean getAppendRows()
  {
    return this.appendRows;
  }
  
  public void setTableColumns(String[] tableColumns)
  {
    this.tableColumns = tableColumns;
  }
  
  public String[] getTableColumns()
  {
    return this.tableColumns;
  }
  
  public void setChart(String chart)
  {
    this.chart = chart;
  }
  
  public String getChart()
  {
    return this.chart;
  }
  
  public void setUpdateLabel(boolean updateLabel)
  {
    this.updateLabel = updateLabel;
  }
  
  public boolean getUpdateLabel()
  {
    return this.updateLabel;
  }
  
  public void setLabel(String label)
  {
    this.label = label;
  }
  
  public String getLabel()
  {
    return this.label;
  }
  
  public void setCAppendRows(boolean cAppendRows)
  {
    this.cAppendRows = cAppendRows;
  }
  
  public boolean getCAppendRows()
  {
    return this.cAppendRows;
  }
  
  public void setChartColumns(String[] chartColumns)
  {
    this.chartColumns = chartColumns;
  }
  
  public String[] getChartColumns()
  {
    return this.chartColumns;
  }
}


/* Location:              /home/bart/bart.maertens@know.bi/Projecten/Drivolution/word-plugin/pentaho_wordwriter/wordwriter.jar!/org/pentaho/di/trans/steps/wordwriter/WordWriterMeta.class
 * Java compiler version: 6 (50.0)
 * JD-Core Version:       0.7.1
 */