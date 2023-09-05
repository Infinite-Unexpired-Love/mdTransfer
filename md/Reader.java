package mdRelated;
import java.io.IOException;
import java.lang.*;
import java.lang.reflect.Array;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Scanner;

public class mdReader {
    private ArrayList<String> fileLines;
    private String filePath;

    public mdReader(String filePath){
        if(filePath.matches(".*(\.md)$"))
            this.filePath=filePath;
        else
            this.filePath=null;
    }
    private boolean readMd(){
        if(filePath==null){
            System.out.println("[-] 文件类型错误，必须为md格式！");
            return false;
        }
        else{
            try{
                Scanner in=new Scanner(Path.of(filePath), StandardCharsets.UTF_8);
                while (in.hasNextLine())
                    fileLines.add(in.nextLine());
                return true;
            }
            catch (IOException e){
                e.printStackTrace();
                fileLines=null;
                return false;
            }
        }

    }
    public ArrayList<String> getText(){
        if(readMd())
            return new ArrayList<>(fileLines);
        else
            return null;
    },
    public ArrayList<>
}
