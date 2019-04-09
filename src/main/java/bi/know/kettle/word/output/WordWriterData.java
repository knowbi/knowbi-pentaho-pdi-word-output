package bi.know.kettle.word.output;

import bi.know.kettle.word.docx.WordDocument;
//import org.apache.commons.vfs.FileObject;
import org.apache.commons.vfs2.FileObject;
import org.pentaho.di.core.row.RowMetaInterface;
import org.pentaho.di.trans.step.BaseStepData;
import org.pentaho.di.trans.step.StepDataInterface;

public class WordWriterData
  extends BaseStepData
  implements StepDataInterface
{
  public RowMetaInterface outputRowMeta;
  public RowMetaInterface inputRowMeta;
  public FileObject file;
  public WordDocument wdoc;
  public int row;
  public int cRow;
}


/* Location:              /home/bart/bart.maertens@know.bi/Projecten/Drivolution/word-plugin/pentaho_wordwriter/wordwriter.jar!/org/pentaho/di/trans/steps/wordwriter/WordWriterData.class
 * Java compiler version: 6 (50.0)
 * JD-Core Version:       0.7.1
 */