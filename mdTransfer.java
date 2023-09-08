import md.*;
import DocxGen.DocxGenerater;
import ClipBoard.ClipBoard;

import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * 1.读取md文件源代码内容
 * 2.解析md文件内容，普通代码直接复制，图片文件需要格外复制二进制流
 * 3.输出所有内容到系统剪切板中
 * @version 1.0.0
 */
public class mdTransfer {
    /**
     *
     * @param args [0]md文件路径;[1]docx文件路径
     */
    public static void main(String[] args){
        if(args.length==0){
            System.out.print("""
                    
                    *********使用说明*********
                    *******Introduction*******
                    
                    java -jar /path/to/mdTransfer.jar arg1 arg2
                    
                    arg1是md文件路径。arg1 is the path to md file.
                    
                    arg2是docx文件目标路径。arg2 is the path to docx file.
                    
                    **************************
                    
                    """);
            return;
        }
        Reader mdReader=new Reader(args[0]);
        mdReader.readMd();
        try{
            DocxGenerater docx=new DocxGenerater(args[1],mdReader);
            docx.getDocx();
        }
        catch (FileNotFoundException e){
            System.out.println("[-] 目标不是docx文件!");
        }
        catch (IOException e){
            e.printStackTrace();
        }

    }
}
