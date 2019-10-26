package ed.synthsys.util.excel;

import java.io.IOException;
import java.nio.file.Path;
import org.apache.poi.hssf.eventusermodel.AbortableHSSFListener;

import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.HSSFUserException;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.RowRecord;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

class XLSDimensionsChecker {

    
    public int[] rowsColsDimensions(Path inFile) throws IOException {
            SheetToSize sizer = new SheetToSize();

            HSSFEventFactory factory = new HSSFEventFactory();
            HSSFRequest request = new HSSFRequest();

            request.addListenerForAllRecords(sizer);
            
            try (POIFSFileSystem fs = new POIFSFileSystem(inFile.toFile(), true)) {
                factory.processWorkbookEvents(request, fs);
            }
            
            return new int[]{sizer.rows, sizer.cols};
    }    


    static class SheetToSize extends AbortableHSSFListener {

        public int rows = 0;
        public int cols = 0;
        public boolean seenSheet = false;
        
        @Override
        public short abortableProcessRecord(Record record) throws HSSFUserException {
            
            if (record.getSid() == BOFRecord.sid) {
			BOFRecord br = (BOFRecord)record;
			if(br.getType() == BOFRecord.TYPE_WORKSHEET) {
				
                        if (seenSheet) return 1;
                        seenSheet = true;                        
                    }
            }
            
            if (record.getSid() == RowRecord.sid) {
                RowRecord row = (RowRecord)record;
                rows = Math.max(rows, row.getRowNumber()+1);
                cols = Math.max(cols, row.getLastCol());
            }
            
            return 0;
        }


        
    }
}