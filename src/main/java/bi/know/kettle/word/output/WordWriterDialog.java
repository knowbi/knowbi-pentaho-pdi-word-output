package bi.know.kettle.word.output;

import bi.know.kettle.word.docx.WordDocument;
import java.util.List;
//import org.apache.commons.vfs.FileObject;
import org.apache.commons.vfs2.FileObject;
import org.eclipse.swt.custom.CCombo;
import org.eclipse.swt.custom.CTabFolder;
import org.eclipse.swt.custom.CTabItem;
import org.eclipse.swt.custom.ScrolledComposite;
import org.eclipse.swt.events.ModifyEvent;
import org.eclipse.swt.events.ModifyListener;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.events.ShellAdapter;
import org.eclipse.swt.events.ShellEvent;
import org.eclipse.swt.layout.FormAttachment;
import org.eclipse.swt.layout.FormData;
import org.eclipse.swt.layout.FormLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Event;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.Group;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Listener;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.TableItem;
import org.eclipse.swt.widgets.Text;
import org.pentaho.di.core.Const;
import org.pentaho.di.core.exception.KettleException;
import org.pentaho.di.core.row.RowMetaInterface;
import org.pentaho.di.core.vfs.KettleVFS;
import org.pentaho.di.i18n.BaseMessages;
import org.pentaho.di.trans.TransMeta;
import org.pentaho.di.trans.step.BaseStepMeta;
import org.pentaho.di.trans.step.StepDialogInterface;
import org.pentaho.di.ui.core.dialog.ErrorDialog;
import org.pentaho.di.ui.core.widget.ColumnInfo;
import org.pentaho.di.ui.core.widget.TableView;
import org.pentaho.di.ui.core.widget.TextVar;
import org.pentaho.di.ui.trans.step.BaseStepDialog;

public class WordWriterDialog
  extends BaseStepDialog
  implements StepDialogInterface
{
  private static Class<?> PKG = WordWriterMeta.class;
  private WordWriterMeta meta;
  private CTabFolder wTabFolder;
  private CTabItem wFileTab;
  private CTabItem wContentTab;
  private Composite wFileComp;
  private Composite wContentComp;
  private TextVar wFilename;
  private CCombo wFunction;
  private Group wReplaceInTextGroup;
  private CCombo wRIPSearchField;
  private CCombo wRIPReplaceByField;
  private Group wInsertInTableGroup;
  private Group wInsertInTableFieldsGroup;
  private CCombo wIITTable;
  private Text wIITStartRow;
  private Text wIITStartColumn;
  private Button wIITAppendRows;
  private TableView wIITFields;
  private Group wInsertInChartGroup;
  private Group wInsertInChartFieldsGroup;
  private CCombo wIICChart;
  private CCombo wIICLabelField;
  private Button wIICUpdateLabel;
  private Button wIICAppendRows;
  private TableView wIICFields;
  private Button wOK;
  private Button wCancel;
  
  public WordWriterDialog(Shell parent, Object in, TransMeta transMeta, String sname)
  {
    super(parent, (BaseStepMeta)in, transMeta, sname);
    this.meta = ((WordWriterMeta)in);
  }
  
  public String open()
  {
    Shell parent = getParent();
    Display display = parent.getDisplay();
    
    this.shell = new Shell(parent, 3312);
    this.props.setLook(this.shell);
    setShellImage(this.shell, this.meta);
    
    this.changed = this.meta.hasChanged();
    
    ModifyListener lsMod = new ModifyListener()
    {
      public void modifyText(ModifyEvent e)
      {
        WordWriterDialog.this.meta.setChanged();
      }
    };
    String[] fieldNames = null;
    try
    {
      RowMetaInterface r = this.transMeta.getPrevStepFields(this.stepname);
      if (r != null) {
        fieldNames = r.getFieldNames();
      }
    }
    catch (KettleException ke)
    {
      new ErrorDialog(this.shell, BaseMessages.getString(PKG, "", new String[0]), BaseMessages.getString(PKG, "", new String[0]), ke);
    }
    FormLayout formLayout = new FormLayout();
    formLayout.marginWidth = 5;
    formLayout.marginHeight = 5;
    
    this.shell.setLayout(formLayout);
    this.shell.setText(BaseMessages.getString(PKG, "WordWriter.Shell.Title", new String[0]));
    
    int middle = this.props.getMiddlePct();
    int margin = 4;
    
    this.wlStepname = new Label(this.shell, 131072);
    this.wlStepname.setText(BaseMessages.getString(PKG, "System.Label.StepName", new String[0]));
    this.props.setLook(this.wlStepname);
    this.fdlStepname = new FormData();
    this.fdlStepname.left = new FormAttachment(0, 0);
    this.fdlStepname.right = new FormAttachment(middle, -margin);
    this.fdlStepname.top = new FormAttachment(0, margin);
    this.wlStepname.setLayoutData(this.fdlStepname);
    
    this.wStepname = new Text(this.shell, 18436);
    this.wStepname.setText(this.stepname);
    this.props.setLook(this.wStepname);
    this.wStepname.addModifyListener(lsMod);
    this.fdStepname = new FormData();
    this.fdStepname.left = new FormAttachment(middle, 0);
    this.fdStepname.top = new FormAttachment(0, margin);
    this.fdStepname.right = new FormAttachment(100, 0);
    this.wStepname.setLayoutData(this.fdStepname);
    
    ScrolledComposite sc = new ScrolledComposite(this.shell, 768);
    
    this.wTabFolder = new CTabFolder(sc, 2048);
    this.props.setLook(this.wTabFolder, 5);
    
    this.wFileTab = new CTabItem(this.wTabFolder, 0);
    this.wFileTab.setText(BaseMessages.getString(PKG, "WordWriter.FileTab.TabTitle", new String[0]));
    
    this.wFileComp = new Composite(this.wTabFolder, 0);
    this.props.setLook(this.wFileComp);
    
    FormLayout fileLayout = new FormLayout();
    fileLayout.marginWidth = 3;
    fileLayout.marginHeight = 3;
    this.wFileComp.setLayout(fileLayout);
    
    Group fileGroup = new Group(this.wFileComp, 32);
    this.props.setLook(fileGroup);
    fileGroup.setText(BaseMessages.getString(PKG, "WordWriter.FileGroup.Label", new String[0]));
    
    FormLayout fileGroupLayout = new FormLayout();
    fileGroupLayout.marginWidth = 10;
    fileGroupLayout.marginHeight = 10;
    fileGroup.setLayout(fileGroupLayout);
    
    Label wlFilename = new Label(fileGroup, 131072);
    wlFilename.setText(BaseMessages.getString(PKG, "WordWriter.Filename.Label", new String[0]));
    this.props.setLook(wlFilename);
    FormData fdlFilename = new FormData();
    fdlFilename.left = new FormAttachment(0, 0);
    fdlFilename.top = new FormAttachment(0, margin);
    fdlFilename.right = new FormAttachment(middle, -margin);
    wlFilename.setLayoutData(fdlFilename);
    
    Button wbFilename = new Button(fileGroup, 16777224);
    this.props.setLook(wbFilename);
    wbFilename.setText(BaseMessages.getString(PKG, "System.Button.Browse", new String[0]));
    FormData fdbFilename = new FormData();
    fdbFilename.right = new FormAttachment(100, 0);
    fdbFilename.top = new FormAttachment(0, 0);
    wbFilename.setLayoutData(fdbFilename);
    
    this.wFilename = new TextVar(this.transMeta, fileGroup, 18436);
    this.props.setLook(this.wFilename);
    this.wFilename.addModifyListener(lsMod);
    this.wFilename.setToolTipText(BaseMessages.getString(PKG, "WordWriter.Filename.Tooltip", new String[0]));
    FormData fdFilename = new FormData();
    fdFilename.left = new FormAttachment(middle, 0);
    fdFilename.top = new FormAttachment(0, margin);
    fdFilename.right = new FormAttachment(wbFilename, -margin);
    this.wFilename.setLayoutData(fdFilename);
    
    FormData fsFileGroup = new FormData();
    fsFileGroup.left = new FormAttachment(0, margin);
    fsFileGroup.top = new FormAttachment(0, margin);
    fsFileGroup.right = new FormAttachment(100, -margin);
    fileGroup.setLayoutData(fsFileGroup);
    
    Group functionGroup = new Group(this.wFileComp, 32);
    this.props.setLook(functionGroup);
    functionGroup.setText(BaseMessages.getString(PKG, "WordWriter.FunctionGroup.Label", new String[0]));
    
    FormLayout functionGroupLayout = new FormLayout();
    functionGroupLayout.marginWidth = 10;
    functionGroupLayout.marginHeight = 10;
    functionGroup.setLayout(functionGroupLayout);
    
    Label wlFunction = new Label(functionGroup, 131072);
    wlFunction.setText(BaseMessages.getString(PKG, "WordWriter.Function.Label", new String[0]));
    this.props.setLook(wlFunction);
    FormData fdlFunction = new FormData();
    fdlFunction.left = new FormAttachment(0, 0);
    fdlFunction.top = new FormAttachment(0, margin);
    fdlFunction.right = new FormAttachment(middle, -margin);
    wlFunction.setLayoutData(fdlFunction);
    
    this.wFunction = new CCombo(functionGroup, 131072);
    
    String replaceInTextLabel = BaseMessages.getString(PKG, "WordWriter.Function.ReplaceInText.Label", new String[0]);
    String insertInTableLabel = BaseMessages.getString(PKG, "WordWriter.Function.InsertInTable.Label", new String[0]);
    String insertInChartLabel = BaseMessages.getString(PKG, "WordWriter.Function.InsertInChart.Label", new String[0]);
    String[] functionsLabels = { replaceInTextLabel, insertInTableLabel, insertInChartLabel };
    this.wFunction.setItems(functionsLabels);
    this.wFunction.setData(replaceInTextLabel, "replace_in_text");
    this.wFunction.setData(insertInTableLabel, "insert_in_table");
    this.wFunction.setData(insertInChartLabel, "insert_in_chart");
    
    this.props.setLook(this.wFunction);
    this.wFunction.addModifyListener(lsMod);
    this.wFunction.addSelectionListener(new SelectionAdapter()
    {
      public void widgetSelected(SelectionEvent e)
      {
        WordWriterDialog.this.changeContent();
      }
    });
    this.wFunction.setToolTipText(BaseMessages.getString(PKG, "WordWriter.Function.Tooltip", new String[0]));
    
    FormData fdFunction = new FormData();
    fdFunction.left = new FormAttachment(middle, 0);
    fdFunction.top = new FormAttachment(0, margin);
    fdFunction.right = new FormAttachment(100, 0);
    this.wFunction.setLayoutData(fdFunction);
    
    FormData fsFunctionGroup = new FormData();
    fsFunctionGroup.left = new FormAttachment(0, margin);
    fsFunctionGroup.top = new FormAttachment(fileGroup, margin);
    fsFunctionGroup.right = new FormAttachment(100, -margin);
    functionGroup.setLayoutData(fsFunctionGroup);
    
    FormData fdFileComp = new FormData();
    fdFileComp.left = new FormAttachment(0, 0);
    fdFileComp.top = new FormAttachment(0, 0);
    fdFileComp.right = new FormAttachment(100, 0);
    fdFileComp.bottom = new FormAttachment(100, 0);
    this.wFileComp.setLayoutData(fdFileComp);
    
    this.wFileComp.layout();
    this.wFileTab.setControl(this.wFileComp);
    
    this.wContentTab = new CTabItem(this.wTabFolder, 0);
    this.wContentTab.setText(BaseMessages.getString(PKG, "WordWriter.ContentTab.TabTitle", new String[0]));
    
    this.wContentComp = new Composite(this.wTabFolder, 0);
    this.props.setLook(this.wContentComp);
    
    FormLayout contentLayout = new FormLayout();
    contentLayout.marginWidth = 3;
    contentLayout.marginHeight = 3;
    this.wContentComp.setLayout(contentLayout);
    
    this.wReplaceInTextGroup = new Group(this.wContentComp, 32);
    this.props.setLook(this.wReplaceInTextGroup);
    this.wReplaceInTextGroup.setVisible(false);
    this.wReplaceInTextGroup.setText(BaseMessages.getString(PKG, "WordWriter.Content.Group.ReplaceInText.Label", new String[0]));
    
    FormLayout replaceInTextGroupLayout = new FormLayout();
    replaceInTextGroupLayout.marginWidth = 10;
    replaceInTextGroupLayout.marginHeight = 10;
    this.wReplaceInTextGroup.setLayout(replaceInTextGroupLayout);
    
    Label wlRIPSearchField = new Label(this.wReplaceInTextGroup, 131072);
    wlRIPSearchField.setText(BaseMessages.getString(PKG, "WordWriter.Content.ReplaceInText.SearchField.Label", new String[0]));
    this.props.setLook(wlRIPSearchField);
    FormData fdlRIPSearchField = new FormData();
    fdlRIPSearchField.left = new FormAttachment(0, 0);
    fdlRIPSearchField.top = new FormAttachment(0, margin);
    fdlRIPSearchField.right = new FormAttachment(middle, -margin);
    wlRIPSearchField.setLayoutData(fdlRIPSearchField);
    
    this.wRIPSearchField = new CCombo(this.wReplaceInTextGroup, 131072);
    this.props.setLook(this.wRIPSearchField);
    this.wRIPSearchField.addModifyListener(lsMod);
    FormData fdRIPSearchField = new FormData();
    fdRIPSearchField.left = new FormAttachment(middle, 0);
    fdRIPSearchField.top = new FormAttachment(0, margin);
    fdRIPSearchField.right = new FormAttachment(100, 0);
    this.wRIPSearchField.setLayoutData(fdRIPSearchField);
    if (fieldNames != null) {
      this.wRIPSearchField.setItems(fieldNames);
    }
    Label wlRIPReplaceByField = new Label(this.wReplaceInTextGroup, 131072);
    wlRIPReplaceByField.setText(BaseMessages.getString(PKG, "WordWriter.Content.ReplaceInText.ReplaceByField.Label", new String[0]));
    this.props.setLook(wlRIPReplaceByField);
    FormData fdlRIPReplaceByField = new FormData();
    fdlRIPReplaceByField.left = new FormAttachment(0, 0);
    fdlRIPReplaceByField.top = new FormAttachment(this.wRIPSearchField, margin);
    fdlRIPReplaceByField.right = new FormAttachment(middle, -margin);
    wlRIPReplaceByField.setLayoutData(fdlRIPReplaceByField);
    
    this.wRIPReplaceByField = new CCombo(this.wReplaceInTextGroup, 131072);
    this.props.setLook(this.wRIPReplaceByField);
    this.wRIPReplaceByField.addModifyListener(lsMod);
    FormData fdRIPReplaceByField = new FormData();
    fdRIPReplaceByField.left = new FormAttachment(middle, 0);
    fdRIPReplaceByField.top = new FormAttachment(this.wRIPSearchField, margin);
    fdRIPReplaceByField.right = new FormAttachment(100, 0);
    this.wRIPReplaceByField.setLayoutData(fdRIPReplaceByField);
    if (fieldNames != null) {
      this.wRIPReplaceByField.setItems(fieldNames);
    }
    FormData fsReplaceInTextGroup = new FormData();
    fsReplaceInTextGroup.left = new FormAttachment(0, margin);
    fsReplaceInTextGroup.top = new FormAttachment(0, margin);
    fsReplaceInTextGroup.right = new FormAttachment(100, -margin);
    this.wReplaceInTextGroup.setLayoutData(fsReplaceInTextGroup);
    
    this.wInsertInTableGroup = new Group(this.wContentComp, 32);
    this.props.setLook(this.wInsertInTableGroup);
    this.wInsertInTableGroup.setVisible(false);
    this.wInsertInTableGroup.setText(BaseMessages.getString(PKG, "WordWriter.Content.Group.InsertInTable.Label", new String[0]));
    
    FormLayout insertInTableGroupLayout = new FormLayout();
    insertInTableGroupLayout.marginWidth = 10;
    insertInTableGroupLayout.marginHeight = 10;
    this.wInsertInTableGroup.setLayout(insertInTableGroupLayout);
    
    Label wlIITTable = new Label(this.wInsertInTableGroup, 131072);
    wlIITTable.setText(BaseMessages.getString(PKG, "WordWriter.Content.InsertInTable.Table.Label", new String[0]));
    this.props.setLook(wlIITTable);
    FormData fdlIITTable = new FormData();
    fdlIITTable.left = new FormAttachment(0, 0);
    fdlIITTable.top = new FormAttachment(0, margin);
    fdlIITTable.right = new FormAttachment(middle, -margin);
    wlIITTable.setLayoutData(fdlIITTable);
    
    Button wbIITTable = new Button(this.wInsertInTableGroup, 16777224);
    this.props.setLook(wbIITTable);
    wbIITTable.setText(BaseMessages.getString(PKG, "WordWriter.Content.InsertInTable.Table.Button", new String[0]));
    FormData fdbIITTable = new FormData();
    fdbIITTable.right = new FormAttachment(100, 0);
    fdbIITTable.top = new FormAttachment(0, 0);
    wbIITTable.setLayoutData(fdbIITTable);
    
    this.wIITTable = new CCombo(this.wInsertInTableGroup, 18436);
    this.props.setLook(this.wIITTable);
    this.wIITTable.addModifyListener(lsMod);
    FormData fdIITTable = new FormData();
    fdIITTable.left = new FormAttachment(middle, 0);
    fdIITTable.top = new FormAttachment(0, margin);
    fdIITTable.right = new FormAttachment(wbIITTable, -margin);
    this.wIITTable.setLayoutData(fdIITTable);
    
    Label wlIITStartRow = new Label(this.wInsertInTableGroup, 131072);
    wlIITStartRow.setText(BaseMessages.getString(PKG, "WordWriter.Content.InsertInTable.StartRow.Label", new String[0]));
    this.props.setLook(wlIITStartRow);
    FormData fdlIITStartRow = new FormData();
    fdlIITStartRow.left = new FormAttachment(0, 0);
    fdlIITStartRow.top = new FormAttachment(wbIITTable, margin);
    fdlIITStartRow.right = new FormAttachment(middle, -margin);
    wlIITStartRow.setLayoutData(fdlIITStartRow);
    
    this.wIITStartRow = new Text(this.wInsertInTableGroup, 18436);
    this.props.setLook(this.wIITStartRow);
    this.wIITStartRow.addModifyListener(lsMod);
    FormData fdIITStartRow = new FormData();
    fdIITStartRow.left = new FormAttachment(middle, 0);
    fdIITStartRow.top = new FormAttachment(wbIITTable, margin);
    fdIITStartRow.right = new FormAttachment(100, 0);
    this.wIITStartRow.setLayoutData(fdIITStartRow);
    
    Label wlIITStartColumn = new Label(this.wInsertInTableGroup, 131072);
    wlIITStartColumn.setText(BaseMessages.getString(PKG, "WordWriter.Content.InsertInTable.StartColumn.Label", new String[0]));
    this.props.setLook(wlIITStartColumn);
    FormData fdlIITStartColumn = new FormData();
    fdlIITStartColumn.left = new FormAttachment(0, 0);
    fdlIITStartColumn.top = new FormAttachment(this.wIITStartRow, margin);
    fdlIITStartColumn.right = new FormAttachment(middle, -margin);
    wlIITStartColumn.setLayoutData(fdlIITStartColumn);
    
    this.wIITStartColumn = new Text(this.wInsertInTableGroup, 18436);
    this.props.setLook(this.wIITStartColumn);
    this.wIITStartColumn.addModifyListener(lsMod);
    FormData fdIITStartColumn = new FormData();
    fdIITStartColumn.left = new FormAttachment(middle, 0);
    fdIITStartColumn.top = new FormAttachment(this.wIITStartRow, margin);
    fdIITStartColumn.right = new FormAttachment(100, 0);
    this.wIITStartColumn.setLayoutData(fdIITStartColumn);
    
    Label wlIITAppendRows = new Label(this.wInsertInTableGroup, 131072);
    wlIITAppendRows.setText(BaseMessages.getString(PKG, "WordWriter.Content.InsertInTable.AppendRows.Label", new String[0]));
    this.props.setLook(wlIITAppendRows);
    FormData fdlIITAppendRows = new FormData();
    fdlIITAppendRows.left = new FormAttachment(0, 0);
    fdlIITAppendRows.top = new FormAttachment(this.wIITStartColumn, margin);
    fdlIITAppendRows.right = new FormAttachment(middle, -margin);
    wlIITAppendRows.setLayoutData(fdlIITAppendRows);
    
    this.wIITAppendRows = new Button(this.wInsertInTableGroup, 32);
    this.props.setLook(this.wIITAppendRows);
    FormData fdIITAppendRows = new FormData();
    fdIITAppendRows.left = new FormAttachment(middle, 0);
    fdIITAppendRows.top = new FormAttachment(this.wIITStartColumn, margin);
    fdIITAppendRows.right = new FormAttachment(100, 0);
    this.wIITAppendRows.setLayoutData(fdIITAppendRows);
    this.wIITAppendRows.addSelectionListener(new SelectionAdapter()
    {
      public void widgetSelected(SelectionEvent e)
      {
        WordWriterDialog.this.meta.setChanged();
      }
    });
    FormData fsInsertInTableGroup = new FormData();
    fsInsertInTableGroup.left = new FormAttachment(0, margin);
    fsInsertInTableGroup.top = new FormAttachment(0, margin);
    fsInsertInTableGroup.right = new FormAttachment(100, -margin);
    this.wInsertInTableGroup.setLayoutData(fsInsertInTableGroup);
    
    this.wInsertInTableFieldsGroup = new Group(this.wContentComp, 32);
    this.props.setLook(this.wInsertInTableFieldsGroup);
    this.wInsertInTableFieldsGroup.setVisible(false);
    this.wInsertInTableFieldsGroup.setText(BaseMessages.getString(PKG, "WordWriter.Content.Group.InsertInTable.Fields.Label", new String[0]));
    
    FormLayout insertInTableFieldsGroupLayout = new FormLayout();
    insertInTableFieldsGroupLayout.marginWidth = 10;
    insertInTableFieldsGroupLayout.marginHeight = 10;
    this.wInsertInTableFieldsGroup.setLayout(insertInTableFieldsGroupLayout);
    
    int fieldRows = this.meta.getTableColumns().length;
    ColumnInfo[] colInf = { new ColumnInfo(BaseMessages.getString(PKG, "WordWriter.Content.InsertInTable.FieldName", new String[0]), 2, fieldNames, false) };
    
    this.wIITFields = new TableView(this.transMeta, this.wInsertInTableFieldsGroup, 67586, colInf, fieldRows, lsMod, this.props);
    
    FormData fdIITFields = new FormData();
    fdIITFields.left = new FormAttachment(0, 0);
    fdIITFields.top = new FormAttachment(0, 0);
    fdIITFields.right = new FormAttachment(100, 0);
    fdIITFields.bottom = new FormAttachment(100, -margin);
    this.wIITFields.setLayoutData(fdIITFields);
    this.wIITFields.addModifyListener(lsMod);
    
    FormData fdInsertInTableFieldsGroup = new FormData();
    fdInsertInTableFieldsGroup.left = new FormAttachment(0, margin);
    fdInsertInTableFieldsGroup.top = new FormAttachment(this.wInsertInTableGroup, margin);
    fdInsertInTableFieldsGroup.bottom = new FormAttachment(100, 0);
    fdInsertInTableFieldsGroup.right = new FormAttachment(100, -margin);
    this.wInsertInTableFieldsGroup.setLayoutData(fdInsertInTableFieldsGroup);
    
    this.wInsertInChartGroup = new Group(this.wContentComp, 32);
    this.props.setLook(this.wInsertInChartGroup);
    this.wInsertInChartGroup.setVisible(false);
    this.wInsertInChartGroup.setText(BaseMessages.getString(PKG, "WordWriter.Content.Group.InsertInChart.Label", new String[0]));
    
    FormLayout insertInChartGroupLayout = new FormLayout();
    insertInChartGroupLayout.marginWidth = 10;
    insertInChartGroupLayout.marginHeight = 10;
    this.wInsertInChartGroup.setLayout(insertInChartGroupLayout);
    
    Label wlIICChart = new Label(this.wInsertInChartGroup, 131072);
    this.props.setLook(wlIICChart);
    wlIICChart.setText(BaseMessages.getString(PKG, "WordWriter.Content.InsertInChart.Chart.Label", new String[0]));
    FormData fdlIICChart = new FormData();
    fdlIICChart.left = new FormAttachment(0, 0);
    fdlIICChart.top = new FormAttachment(0, margin);
    fdlIICChart.right = new FormAttachment(middle, -margin);
    wlIICChart.setLayoutData(fdlIICChart);
    
    Button wbIICChart = new Button(this.wInsertInChartGroup, 16777224);
    this.props.setLook(wbIICChart);
    wbIICChart.setText(BaseMessages.getString(PKG, "WordWriter.Content.InsertInChart.Chart.Button", new String[0]));
    FormData fdbIICChart = new FormData();
    fdbIICChart.right = new FormAttachment(100, 0);
    fdbIICChart.top = new FormAttachment(0, 0);
    wbIICChart.setLayoutData(fdbIICChart);
    
    this.wIICChart = new CCombo(this.wInsertInChartGroup, 18436);
    this.props.setLook(this.wIICChart);
    this.wIICChart.addModifyListener(lsMod);
    FormData fdIICChart = new FormData();
    fdIICChart.left = new FormAttachment(middle, 0);
    fdIICChart.top = new FormAttachment(0, margin);
    fdIICChart.right = new FormAttachment(wbIICChart, -margin);
    this.wIICChart.setLayoutData(fdIICChart);
    
    Label wlIICUpdateLabel = new Label(this.wInsertInChartGroup, 131072);
    wlIICUpdateLabel.setText(BaseMessages.getString(PKG, "WordWriter.Content.InsertInChart.UpdateLabel.Label", new String[0]));
    this.props.setLook(wlIICUpdateLabel);
    FormData fdlIICUpdateLabel = new FormData();
    fdlIICUpdateLabel.left = new FormAttachment(0, 0);
    fdlIICUpdateLabel.top = new FormAttachment(wbIICChart, margin);
    fdlIICUpdateLabel.right = new FormAttachment(middle, -margin);
    wlIICUpdateLabel.setLayoutData(fdlIICUpdateLabel);
    
    this.wIICUpdateLabel = new Button(this.wInsertInChartGroup, 32);
    this.props.setLook(this.wIICUpdateLabel);
    FormData fdIICUpdateLabel = new FormData();
    fdIICUpdateLabel.left = new FormAttachment(middle, 0);
    fdIICUpdateLabel.top = new FormAttachment(wbIICChart, margin);
    fdIICUpdateLabel.right = new FormAttachment(100, 0);
    this.wIICUpdateLabel.setLayoutData(fdIICUpdateLabel);
    this.wIICUpdateLabel.addSelectionListener(new SelectionAdapter()
    {
      public void widgetSelected(SelectionEvent e)
      {
        WordWriterDialog.this.meta.setChanged();
      }
    });
    Label wlIICLabelField = new Label(this.wInsertInChartGroup, 131072);
    wlIICLabelField.setText(BaseMessages.getString(PKG, "WordWriter.Content.InsertInChart.LabelField.Label", new String[0]));
    this.props.setLook(wlIICLabelField);
    FormData fdlIICLabelField = new FormData();
    fdlIICLabelField.left = new FormAttachment(0, 0);
    fdlIICLabelField.top = new FormAttachment(this.wIICUpdateLabel, margin);
    fdlIICLabelField.right = new FormAttachment(middle, -margin);
    wlIICLabelField.setLayoutData(fdlIICLabelField);
    
    this.wIICLabelField = new CCombo(this.wInsertInChartGroup, 18436);
    this.props.setLook(this.wIICLabelField);
    this.wIICLabelField.addModifyListener(lsMod);
    FormData fdIICLabelField = new FormData();
    fdIICLabelField.left = new FormAttachment(middle, 0);
    fdIICLabelField.top = new FormAttachment(this.wIICUpdateLabel, margin);
    fdIICLabelField.right = new FormAttachment(100, 0);
    this.wIICLabelField.setLayoutData(fdIICLabelField);
    if (fieldNames != null) {
      this.wIICLabelField.setItems(fieldNames);
    }
    Label wlIICAppendRows = new Label(this.wInsertInChartGroup, 131072);
    wlIICAppendRows.setText(BaseMessages.getString(PKG, "WordWriter.Content.InsertInChart.AppendRows.Label", new String[0]));
    this.props.setLook(wlIICAppendRows);
    FormData fdlIICAppendRows = new FormData();
    fdlIICAppendRows.left = new FormAttachment(0, 0);
    fdlIICAppendRows.top = new FormAttachment(this.wIICLabelField, margin);
    fdlIICAppendRows.right = new FormAttachment(middle, -margin);
    wlIICAppendRows.setLayoutData(fdlIICAppendRows);
    
    this.wIICAppendRows = new Button(this.wInsertInChartGroup, 32);
    this.props.setLook(this.wIICAppendRows);
    FormData fdIICAppendRows = new FormData();
    fdIICAppendRows.left = new FormAttachment(middle, 0);
    fdIICAppendRows.top = new FormAttachment(this.wIICLabelField, margin);
    fdIICAppendRows.right = new FormAttachment(100, 0);
    this.wIICAppendRows.setLayoutData(fdIICAppendRows);
    this.wIICAppendRows.addSelectionListener(new SelectionAdapter()
    {
      public void widgetSelected(SelectionEvent e)
      {
        WordWriterDialog.this.meta.setChanged();
      }
    });
    FormData fsInsertInChartGroup = new FormData();
    fsInsertInChartGroup.left = new FormAttachment(0, margin);
    fsInsertInChartGroup.top = new FormAttachment(0, margin);
    fsInsertInChartGroup.right = new FormAttachment(100, -margin);
    this.wInsertInChartGroup.setLayoutData(fsInsertInChartGroup);
    
    this.wInsertInChartFieldsGroup = new Group(this.wContentComp, 32);
    this.props.setLook(this.wInsertInChartFieldsGroup);
    this.wInsertInChartFieldsGroup.setVisible(false);
    this.wInsertInChartFieldsGroup.setText(BaseMessages.getString(PKG, "WordWriter.Content.Group.InsertInChart.Fields.Label", new String[0]));
    
    FormLayout insertInChartFieldsGroupLayout = new FormLayout();
    insertInChartFieldsGroupLayout.marginWidth = 10;
    insertInChartFieldsGroupLayout.marginHeight = 10;
    this.wInsertInChartFieldsGroup.setLayout(insertInChartGroupLayout);
    
    int cFieldRows = this.meta.getChartColumns().length;
    ColumnInfo[] cColInf = { new ColumnInfo(BaseMessages.getString(PKG, "WordWriter.Content.InsertInChart.FieldName", new String[0]), 2, fieldNames, false) };
    
    this.wIICFields = new TableView(this.transMeta, this.wInsertInChartFieldsGroup, 67586, cColInf, cFieldRows, lsMod, this.props);
    
    FormData fdIICFields = new FormData();
    fdIICFields.left = new FormAttachment(0, 0);
    fdIICFields.top = new FormAttachment(0, 0);
    fdIICFields.right = new FormAttachment(100, 0);
    fdIICFields.bottom = new FormAttachment(100, -margin);
    this.wIICFields.setLayoutData(fdIICFields);
    this.wIICFields.addModifyListener(lsMod);
    
    FormData fdInsertInChartFieldsGroup = new FormData();
    fdInsertInChartFieldsGroup.left = new FormAttachment(0, margin);
    fdInsertInChartFieldsGroup.top = new FormAttachment(this.wInsertInChartGroup, margin);
    fdInsertInChartFieldsGroup.bottom = new FormAttachment(100, 0);
    fdInsertInChartFieldsGroup.right = new FormAttachment(100, -margin);
    this.wInsertInChartFieldsGroup.setLayoutData(fdInsertInChartFieldsGroup);
    
    FormData fdContentComp = new FormData();
    fdContentComp.left = new FormAttachment(0, 0);
    fdContentComp.top = new FormAttachment(0, 0);
    fdContentComp.right = new FormAttachment(100, 0);
    fdContentComp.bottom = new FormAttachment(100, 0);
    this.wContentComp.setLayoutData(fdContentComp);
    
    this.wContentComp.layout();
    this.wContentTab.setControl(this.wContentComp);
    
    FormData fdTabFolder = new FormData();
    fdTabFolder.left = new FormAttachment(0, 0);
    fdTabFolder.top = new FormAttachment(0, 0);
    fdTabFolder.right = new FormAttachment(100, 0);
    fdTabFolder.bottom = new FormAttachment(100, 0);
    this.wTabFolder.setLayoutData(fdTabFolder);
    
    FormData fdSc = new FormData();
    fdSc.left = new FormAttachment(0, 0);
    fdSc.top = new FormAttachment(this.wStepname, margin);
    fdSc.right = new FormAttachment(100, 0);
    fdSc.bottom = new FormAttachment(100, -50);
    sc.setLayoutData(fdSc);
    
    sc.setContent(this.wTabFolder);
    
    this.wOK = new Button(this.shell, 8);
    this.wOK.setText(BaseMessages.getString(PKG, "System.Button.OK", new String[0]));
    
    this.wCancel = new Button(this.shell, 8);
    this.wCancel.setText(BaseMessages.getString(PKG, "System.Button.Cancel", new String[0]));
    
    setButtonPositions(new Button[] { this.wOK, this.wCancel }, margin, sc);
    
    Listener lsOK = new Listener()
    {
      public void handleEvent(Event e)
      {
        WordWriterDialog.this.ok();
      }
    };
    Listener lsCancel = new Listener()
    {
      public void handleEvent(Event e)
      {
        WordWriterDialog.this.cancel();
      }
    };
    this.wOK.addListener(13, lsOK);
    this.wCancel.addListener(13, lsCancel);
    
    wbFilename.addSelectionListener(new SelectionAdapter()
    {
      public void widgetSelected(SelectionEvent e)
      {
        FileDialog dialog = new FileDialog(WordWriterDialog.this.shell, 8192);
        dialog.setFilterExtensions(new String[] { "*.docx" });
        if (WordWriterDialog.this.wFilename.getText() != null) {
          dialog.setFileName(WordWriterDialog.this.transMeta.environmentSubstitute(WordWriterDialog.this.wFilename.getText()));
        }
        dialog.setFilterNames(new String[] { BaseMessages.getString(WordWriterDialog.PKG, "WordWriter.FormatDOCX.Label", new String[0]) });
        if (dialog.open() != null) {
          WordWriterDialog.this.wFilename.setText(dialog.getFilterPath() + System.getProperty("file.separator") + dialog.getFileName());
        }
      }
    });
    wbIITTable.addSelectionListener(new SelectionAdapter()
    {
      public void widgetSelected(SelectionEvent selectionEvent)
      {
        if ((WordWriterDialog.this.wFilename.getText() == null) || (WordWriterDialog.this.wFilename.getText().isEmpty())) {
          return;
        }
        String filename = WordWriterDialog.this.transMeta.environmentSubstitute(WordWriterDialog.this.wFilename.getText());
        try
        {
          Integer iITTable = null;
          if (WordWriterDialog.this.wIITTable.getText() != null)
          {
            iITTable = (Integer)WordWriterDialog.this.wIITTable.getData(WordWriterDialog.this.wIITTable.getText());
            if ((iITTable == null) && (WordWriterDialog.this.wIITTable.getText().matches("\\d+"))) {
              iITTable = Integer.valueOf(Integer.parseInt(WordWriterDialog.this.wIITTable.getText()));
            }
          }
          FileObject file = KettleVFS.getFileObject(filename, WordWriterDialog.this.transMeta);
          WordDocument doc = new WordDocument(KettleVFS.getInputStream(file));
          int nrOfTables = doc.getNumberOfTables();
          WordWriterDialog.this.wIITTable.removeAll();
          for (int i = 1; i <= nrOfTables; i++)
          {
            String tableContent = doc.getStringTableContent(i);
            if (tableContent.length() > 50) {
              tableContent = tableContent.substring(0, 50);
            }
            String item = i + ": " + tableContent;
            WordWriterDialog.this.wIITTable.add(item);
            WordWriterDialog.this.wIITTable.setData(item, Integer.valueOf(i));
          }
          if ((iITTable != null) && (iITTable.intValue() != 0)) {
            WordWriterDialog.this.wIITTable.select(iITTable.intValue() - 1);
          }
        }
        catch (Exception e)
        {
          new ErrorDialog(WordWriterDialog.this.shell, BaseMessages.getString(WordWriterDialog.PKG, "System.Dialog.Error.Title", new String[0]), BaseMessages.getString(WordWriterDialog.PKG, "WordWriter.Error.OpeningDocument", new String[] { filename, e.toString() }), e);
        }
      }
    });
    wbIICChart.addSelectionListener(new SelectionAdapter()
    {
      public void widgetSelected(SelectionEvent selectionEvent)
      {
        if ((WordWriterDialog.this.wFilename.getText() == null) || (WordWriterDialog.this.wFilename.getText().isEmpty())) {
          return;
        }
        String filename = WordWriterDialog.this.transMeta.environmentSubstitute(WordWriterDialog.this.wFilename.getText());
        try
        {
          String sChartName = null;
          if (WordWriterDialog.this.wIICChart.getText() != null)
          {
            sChartName = (String)WordWriterDialog.this.wIICChart.getData(WordWriterDialog.this.wIICChart.getText());
            if (sChartName == null) {
              sChartName = WordWriterDialog.this.wIICChart.getText();
            }
          }
          FileObject file = KettleVFS.getFileObject(filename, WordWriterDialog.this.transMeta);
          WordDocument doc = new WordDocument(KettleVFS.getInputStream(file));
          List<String> chartNames = doc.getChartNames();
          WordWriterDialog.this.wIICChart.removeAll();
          for (String chartName : chartNames)
          {
            String chart = chartName + ": " + doc.getChart(chartName).getTitle();
            WordWriterDialog.this.wIICChart.add(chart);
            WordWriterDialog.this.wIICChart.setData(chart, chartName);
          }
          int i = chartNames.indexOf(sChartName);
          if (i != -1) {
            WordWriterDialog.this.wIICChart.select(i);
          }
        }
        catch (Exception e)
        {
          new ErrorDialog(WordWriterDialog.this.shell, BaseMessages.getString(WordWriterDialog.PKG, "System.Dialog.Error.Title", new String[0]), BaseMessages.getString(WordWriterDialog.PKG, "WordWriter.Error.OpeningDocument", new String[] { filename, e.toString() }), e);
        }
      }
    });
    this.wIICUpdateLabel.addSelectionListener(new SelectionAdapter()
    {
      public void widgetSelected(SelectionEvent e)
      {
        WordWriterDialog.this.toggleUpdateLabel();
      }
    });
    this.shell.addShellListener(new ShellAdapter()
    {
      public void shellClosed(ShellEvent e)
      {
        WordWriterDialog.this.cancel();
      }
    });
    this.wTabFolder.setSelection(0);
    
    getData();
    toggleUpdateLabel();
    changeContent();
    
    sc.setMinSize(this.wTabFolder.computeSize(-1, -1));
    sc.setExpandHorizontal(true);
    sc.setExpandVertical(true);
    
    this.meta.setChanged(this.changed);
    
    setSize(this.shell, 600, 600, true);
    
    this.shell.open();
    while (!this.shell.isDisposed()) {
      if (!display.readAndDispatch()) {
        display.sleep();
      }
    }
    return this.stepname;
  }
  
  private void getData()
  {
    if (this.meta.getFilename() != null) {
      this.wFilename.setText(this.meta.getFilename());
    }
    this.wFunction.setText(this.meta.getFunction());
    
    this.wRIPSearchField.setText(this.meta.getSearch());
    this.wRIPReplaceByField.setText(this.meta.getReplaceBy());
    
    Integer table = Integer.valueOf(this.meta.getTable());
    if (table.intValue() != 0)
    {
      this.wIITTable.add(table.toString());
      this.wIITTable.setData(table.toString(), table);
      this.wIITTable.select(0);
    }
    this.wIITStartRow.setText(Integer.toString(this.meta.getStartRow()));
    this.wIITStartColumn.setText(Integer.toString(this.meta.getStartColumn()));
    
    this.wIITAppendRows.setSelection(this.meta.getAppendRows());
    for (int i = 0; i < this.meta.getTableColumns().length; i++)
    {
      TableItem item = this.wIITFields.table.getItem(i);
      if (this.meta.getTableColumns()[i] != null) {
        item.setText(1, this.meta.getTableColumns()[i]);
      }
    }
    String chart = this.meta.getChart();
    if ((chart != null) && (!chart.isEmpty()))
    {
      this.wIICChart.add(chart);
      this.wIICChart.setData(chart, chart);
      this.wIICChart.select(0);
    }
    this.wIICUpdateLabel.setSelection(this.meta.getUpdateLabel());
    
    String label = this.meta.getLabel();
    if ((label != null) && (!label.isEmpty())) {
      this.wIICLabelField.setText(label);
    }
    this.wIICAppendRows.setSelection(this.meta.getCAppendRows());
    for (int i = 0; i < this.meta.getChartColumns().length; i++)
    {
      TableItem item = this.wIICFields.table.getItem(i);
      if (this.meta.getChartColumns()[i] != null) {
        item.setText(1, this.meta.getChartColumns()[i]);
      }
    }
  }
  
  private void toggleUpdateLabel()
  {
    if (this.wIICUpdateLabel.getSelection())
    {
      this.wIICLabelField.setEnabled(true);
      this.wIICAppendRows.setEnabled(true);
    }
    else
    {
      this.wIICLabelField.setEnabled(false);
      this.wIICAppendRows.setEnabled(false);
    }
  }
  
  private void changeContent()
  {
    String function = (String)this.wFunction.getData(this.wFunction.getText());
    if (function == null) {
      function = this.wFunction.getText();
    }
    if (function.equals("replace_in_text"))
    {
      this.wFunction.select(0);
      this.wReplaceInTextGroup.setVisible(true);
      this.wInsertInTableGroup.setVisible(false);
      this.wInsertInTableFieldsGroup.setVisible(false);
      this.wInsertInChartGroup.setVisible(false);
      this.wInsertInChartFieldsGroup.setVisible(false);
    }
    else if (function.equals("insert_in_table"))
    {
      this.wFunction.select(1);
      this.wReplaceInTextGroup.setVisible(false);
      this.wInsertInTableGroup.setVisible(true);
      this.wInsertInTableFieldsGroup.setVisible(true);
      this.wInsertInChartGroup.setVisible(false);
      this.wInsertInChartFieldsGroup.setVisible(false);
    }
    else if (function.equals("insert_in_chart"))
    {
      this.wFunction.select(2);
      this.wReplaceInTextGroup.setVisible(false);
      this.wInsertInTableGroup.setVisible(false);
      this.wInsertInTableFieldsGroup.setVisible(false);
      this.wInsertInChartGroup.setVisible(true);
      this.wInsertInChartFieldsGroup.setVisible(true);
    }
    else
    {
      this.wFunction.select(0);
      this.wReplaceInTextGroup.setVisible(true);
      this.wInsertInTableGroup.setVisible(false);
      this.wInsertInTableFieldsGroup.setVisible(false);
      this.wInsertInChartGroup.setVisible(false);
      this.wInsertInChartFieldsGroup.setVisible(false);
    }
  }
  
  private void cancel()
  {
    this.stepname = null;
    this.meta.setChanged(this.backupChanged);
    dispose();
  }
  
  private void ok()
  {
    if (Const.isEmpty(this.wStepname.getText())) {
      return;
    }
    this.stepname = this.wStepname.getText();
    
    getInfo(this.meta);
    dispose();
  }
  
  private void getInfo(WordWriterMeta wwm)
  {
    wwm.setFilename(this.wFilename.getText());
    
    String function = (String)this.wFunction.getData(this.wFunction.getText());
    if (function == null) {
      function = this.wFunction.getText();
    }
    wwm.setFunction(function);
    
    wwm.setSearch(this.wRIPSearchField.getText());
    wwm.setReplaceBy(this.wRIPReplaceByField.getText());
    if (this.wIITTable.getText() != null)
    {
      Integer table = (Integer)this.wIITTable.getData(this.wIITTable.getText());
      if ((table == null) && (this.wIITTable.getText().matches("\\d+"))) {
        table = Integer.valueOf(Integer.parseInt(this.wIITTable.getText()));
      }
      if (table == null) {
        wwm.setTable(0);
      } else {
        wwm.setTable(table.intValue());
      }
    }
    if ((this.wIITStartRow.getText() != null) && (this.wIITStartRow.getText().matches("\\d+")))
    {
      int startRow = Integer.parseInt(this.wIITStartRow.getText());
      if (startRow > 0) {
        wwm.setStartRow(startRow);
      }
    }
    if ((this.wIITStartColumn.getText() != null) && (this.wIITStartColumn.getText().matches("\\d+")))
    {
      int startColumn = Integer.parseInt(this.wIITStartColumn.getText());
      if (startColumn > 0) {
        wwm.setStartColumn(startColumn);
      }
    }
    wwm.setAppendRows(this.wIITAppendRows.getSelection());
    
    int nrfields = this.wIITFields.nrNonEmpty();
    String[] tableColumns = new String[nrfields];
    for (int i = 0; i < nrfields; i++)
    {
      TableItem item = this.wIITFields.getNonEmpty(i);
      tableColumns[i] = item.getText(1);
    }
    wwm.setTableColumns(tableColumns);
    if (this.wIICChart.getText() != null)
    {
      String chart = (String)this.wIICChart.getData(this.wIICChart.getText());
      if (chart == null) {
        chart = this.wIICChart.getText();
      }
      if (chart == null) {
        wwm.setChart("");
      } else {
        wwm.setChart(chart);
      }
    }
    wwm.setUpdateLabel(this.wIICUpdateLabel.getSelection());
    wwm.setLabel(this.wIICLabelField.getText());
    wwm.setCAppendRows(this.wIICAppendRows.getSelection());
    
    nrfields = this.wIICFields.nrNonEmpty();
    String[] chartColumns = new String[nrfields];
    for (int i = 0; i < nrfields; i++)
    {
      TableItem item = this.wIICFields.getNonEmpty(i);
      chartColumns[i] = item.getText(1);
    }
    wwm.setChartColumns(chartColumns);
  }
}


/* Location:              /home/bart/bart.maertens@know.bi/Projecten/Drivolution/word-plugin/pentaho_wordwriter/wordwriter.jar!/org/pentaho/di/trans/steps/wordwriter/WordWriterDialog.class
 * Java compiler version: 6 (50.0)
 * JD-Core Version:       0.7.1
 */