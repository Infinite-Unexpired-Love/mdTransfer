import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.lang.*;
import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.*;
import ClipBoard.ClipBoard;
/**
 *
 * 目标：
 * 1.读取md文件源代码内容
 * 2.解析md文件内容，普通代码直接复制，图片文件需要格外复制二进制流
 * 3.输出所有内容到word文件中
 *
 * @version 1.0.0
 */
public class Main {
    /**
     *
     * @param args [0]md文件的地址,[1]图片文件的目录
     */
    public static void main(String[] args) {
        try {
//            Scanner fileIn = new Scanner(Path.of(args[0]), StandardCharsets.UTF_8);
//            ArrayList<String> lines=new ArrayList<String>();
//            while (fileIn.hasNextLine())
//                lines.add(fileIn.nextLine());
//            for(var line : lines){
//                System.out.println(line);
//            }
            FileInputStream img=new FileInputStream(args[0]);
            byte[] data=new byte[img.available()];
            img.read(data);
            img.close();
//            for(byte cell : data){
//                System.out.printf("%2x",cell);
//            }
            System.out.printf("%x ",data[0]);
            InputStream in = new ByteArrayInputStream(data);
            BufferedImage bImageFromConvert = ImageIO.read(in);
            ClipBoard.setClipboardImage(bImageFromConvert);
//            FileOutputStream fwriter=new FileOutputStream(new String("./tmp.png"));
//            fwriter.write(data);
//            fwriter.close();
        }
        catch(IOException e){
            e.printStackTrace();

        }

    }
}