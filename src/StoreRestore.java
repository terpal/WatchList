import java.io.*;

/**
 * Created by terpal on 15/4/2017.
 * Stores and Restores permanent variables
 */

public class StoreRestore {
    /**
     * Saves the Last Used File Path
     */
    public void Serialize(String path){

        try {
            FileOutputStream fileOut =
                    new FileOutputStream("lastUsedFile.ser");
            ObjectOutputStream out = new ObjectOutputStream(fileOut);
            out.writeObject(path);
            out.close();
            fileOut.close();
        }catch(IOException i) {
            i.printStackTrace();
        }
    }

    /**
     *  Restores the Last Used File Path
     */
    public String Deserialize() {
        String path = null;
        try {
            FileInputStream fileIn = null;
            fileIn = new FileInputStream("lastUsedFile.ser");
            ObjectInputStream in = new ObjectInputStream(fileIn);
            path = (String) in.readObject();
            in.close();
            fileIn.close();
        } catch (IOException i) {
            i.printStackTrace();
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }
        return path;
    }

}
