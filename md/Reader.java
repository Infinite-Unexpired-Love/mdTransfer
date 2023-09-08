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

    private ArrayList<Integer> imgLines;
    public Reader(String filePath){
        fileLines=new ArrayList<String>();
        imgURLs=new ArrayList<String>();
        imgLines=new ArrayList<Integer>();
        if(filePath.matches(".*(\\.md)$"))
            this.filePath=filePath;
        else
            this.filePath=null;
    }
    private String addSlashes(String origin){
        return origin.replaceAll("\\\\","\\\\");
    }
    public boolean readMd(){
        if(filePath==null){
            System.out.println("[-] 文件类型错误，必须为md格式！");
            return false;
        }
        else{
            try{
                Scanner in=new Scanner(Path.of(filePath), StandardCharsets.UTF_8);
                int lineNum=-1;
                while (in.hasNextLine()){
                    lineNum++;
                    String line=in.nextLine();
                    //转义特殊字符
                    line=line.replaceAll("&","&amp;");
                    line=line.replaceAll("<","&lt;");
                    line=line.replaceAll(">","&gt;");
                    if(line.matches("!\\[.*\\]\\(.*\\)")){
                        imgLines.add(lineNum);
                        int begin=line.indexOf('(')+1;
                        int end=line.indexOf(')');
                        imgURLs.add(addSlashes(line.substring(begin,end)));
                    }
                    fileLines.add(line);
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
    public StringBuffer getText(){
        if(readMd()){
            StringBuffer mdData=new StringBuffer();
            for(var line : fileLines){
                mdData.append(line+'\n');
            }
            return mdData;
        }

        else
            return null;
    }
    public ArrayList<String> getFileLines(){return fileLines;}

    public ArrayList<String> getImgURLs(){
        return imgURLs;
    }

    public ArrayList<Integer> getImgLines(){
        return imgLines;
    }
}
