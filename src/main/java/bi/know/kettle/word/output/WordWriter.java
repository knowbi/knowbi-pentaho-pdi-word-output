package bi.know.kettle.word.output;

import bi.know.kettle.word.docx.WordDocument;
import java.io.BufferedOutputStream;
//import org.apache.commons.vfs2.FileObject;
import org.pentaho.di.core.exception.KettleException;
import org.pentaho.di.core.vfs.KettleVFS;
import org.pentaho.di.i18n.BaseMessages;
import org.pentaho.di.trans.Trans;
import org.pentaho.di.trans.TransMeta;
import org.pentaho.di.trans.step.BaseStep;
import org.pentaho.di.trans.step.StepDataInterface;
import org.pentaho.di.trans.step.StepInterface;
import org.pentaho.di.trans.step.StepMeta;
import org.pentaho.di.trans.step.StepMetaInterface;

public class WordWriter
  extends BaseStep
  implements StepInterface
{
  private static Class<?> PKG = bi.know.kettle.word.output.WordWriterMeta.class;
  private bi.know.kettle.word.output.WordWriterData data;
  private bi.know.kettle.word.output.WordWriterMeta meta;
  
  public WordWriter(StepMeta stepMeta, StepDataInterface stepDataInterface, int copyNr, TransMeta transMeta, Trans trans)
  {
    super(stepMeta, stepDataInterface, copyNr, transMeta, trans);
  }
  
  public boolean init(StepMetaInterface smi, StepDataInterface sdi)
  {
    this.meta = ((WordWriterMeta)smi);
    this.data = ((WordWriterData)sdi);
    
    return super.init(smi, sdi);
  }
  
  public boolean processRow(StepMetaInterface smi, StepDataInterface sdi)
    throws KettleException
  {
    Object[] r = getRow();
    if ((this.first) && (r != null))
    {
      this.first = false;
      try
      {
        prepareOutputFile();
      }
      catch (KettleException e)
      {
        logError("Couldn't prepare output file " + environmentSubstitute(this.meta.getFilename()));
        setErrors(1L);
        stopAll();
      }
      this.data.outputRowMeta = getInputRowMeta().clone();
      this.data.inputRowMeta = getInputRowMeta().clone();
      this.data.row = (this.meta.getStartRow() > 0 ? this.meta.getStartRow() : 0);
      this.data.cRow = 1;
    }
    if (r != null)
    {
      if (this.data.wdoc != null) {
        if ("replace_in_text".equals(this.meta.getFunction())) {
          replaceInText(r);
        } else if ("insert_in_table".equals(this.meta.getFunction())) {
          insertInTable(r);
        } else if ("insert_in_chart".equals(this.meta.getFunction())) {
          insertInChart(r);
        }
      }
      incrementLinesOutput();
      putRow(this.data.outputRowMeta, r);
      if ((checkFeedback(getLinesRead())) && 
        (this.log.isBasic())) {
        logBasic("Linenr " + getLinesOutput());
      }
      return true;
    }
    if (this.data.wdoc != null) {
      closeOutputFile();
    }
    setOutputDone();
    clearDocumentMem();
    
    return false;
  }
  
  public void dispose(StepMetaInterface smi, StepDataInterface sdi)
  {
    this.meta = ((WordWriterMeta)smi);
    this.data = ((WordWriterData)sdi);
    
    clearDocumentMem();
    
    super.dispose(smi, sdi);
  }
  
  private void prepareOutputFile()
    throws KettleException
  {
    String fileName = this.meta.buildFilename(this);
    try
    {
      this.data.file = KettleVFS.getFileObject(fileName, getTransMeta());
      if (!this.data.file.exists())
      {
        if (this.log.isBasic()) {
          logBasic(BaseMessages.getString(PKG, "WordWriter.Log.CouldNotFindFile", new String[] { fileName }));
        }
        throw new KettleException("Could not find file " + fileName);
      }
      this.data.wdoc = new WordDocument(KettleVFS.getInputStream(this.data.file));
    }
    catch (Exception e)
    {
      logError("Error opening file", e);
      setErrors(1L);
      throw new KettleException(e);
    }
  }
  
  private void closeOutputFile()
    throws KettleException
  {
    try
    {
      BufferedOutputStream out = new BufferedOutputStream(KettleVFS.getOutputStream(this.data.file, false));
      this.data.wdoc.write(out);
      out.close();
    }
    catch (Exception e)
    {
      throw new KettleException(e);
    }
  }
  
  private void clearDocumentMem()
  {
    this.data.file = null;
    this.data.wdoc = null;
  }
  
  private synchronized void replaceInText(Object[] r)
    throws KettleException
  {
    String search = getInputRowMeta().getString(r, getInputRowMeta().indexOfValue(this.meta.getSearch()));
    String replaceBy = getInputRowMeta().getString(r, getInputRowMeta().indexOfValue(this.meta.getReplaceBy()));
    this.data.wdoc.replaceInText(search, replaceBy);
  }
  
  private synchronized void insertInTable(Object[] r)
    throws KettleException
  {
    int nrcolumns = this.meta.getTableColumns().length;
    String[] insert = new String[nrcolumns];
    for (int i = 0; i < nrcolumns; i++) {
      insert[i] = getInputRowMeta().getString(r, getInputRowMeta().indexOfValue(this.meta.getTableColumns()[i]));
    }
    this.data.wdoc.insertInTable(this.meta.getTable(), this.data.row, this.meta.getStartColumn(), insert, this.meta.getAppendRows());
    
    this.data.row += 1;
  }
  
  private synchronized void insertInChart(Object[] r)
    throws KettleException
  {
    String label = null;
    if (this.meta.getLabel() != null) {
      label = getInputRowMeta().getString(r, getInputRowMeta().indexOfValue(this.meta.getLabel()));
    }
    int nrcolumns = this.meta.getChartColumns().length;
    Number[] insert = new Number[nrcolumns];
    for (int i = 0; i < nrcolumns; i++) {
      insert[i] = getInputRowMeta().getNumber(r, getInputRowMeta().indexOfValue(this.meta.getChartColumns()[i]));
    }
    this.data.wdoc.getChart(this.meta.getChart()).insertRow(this.data.cRow, insert, this.meta.getUpdateLabel(), label, this.meta.getCAppendRows());
    
    this.data.cRow += 1;
  }
}


/* Location:              /home/bart/bart.maertens@know.bi/Projecten/Drivolution/word-plugin/pentaho_wordwriter/wordwriter.jar!/org/pentaho/di/trans/steps/wordwriter/WordWriter.class
 * Java compiler version: 6 (50.0)
 * JD-Core Version:       0.7.1
 */