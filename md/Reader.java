package md;
import java.io.IOException;
import java.lang.*;
import java.lang.reflect.Array;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Scanner;

public class Reader {
    private ArrayList<String> fileLines;
    private final String filePath;
    private ArrayList<String> imgURLs;
    public Reader(String filePath){
        if(filePath.matches(".*(\\.md)$"))
            this.filePath=filePath;
        else
            this.filePath=null;
    }
    private String addslashes(String origin){
        return origin.replaceAll("\\\\","\\\\");
    }
    private boolean readMd(){
        if(filePath==null){
            System.out.println("[-] 文件类型错误，必须为md格式！");
            return false;
        }
        else{
            try{
                Scanner in=new Scanner(Path.of(filePath), StandardCharsets.UTF_8);
                while (in.hasNextLine()){
                    String line=in.nextLine();

                    if(line.matches("!\\[image-\\d*\\]\\(.*\\)")){
                        int begin=line.indexOf('(');
                        int end=line.indexOf(')');
                        imgURLs.add(addslashes(line.substring(begin+1,end)));
                    }
                    fileLines.add(in.nextLine());
                }
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
    }
    public ArrayList<byte[]> getImages(){

    }
}
