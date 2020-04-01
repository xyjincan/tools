import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.lang.reflect.Field;

public class Main {
	
	public static void main(String[] args) throws Exception {
		int[] keys = new int[args.length];
		for(int i=0;i<args.length;i++) {
			keys[i] = getKeyValue(args[i]);
		}
		PressKey pk = new PressKey();
		pk.Press(keys);
	}
	
	public static String formatKey(String ikey) {
		ikey = ikey.toUpperCase();
		// java.awt.event.KeyEvent  public static final int VK_CONTROL = 0x11;
		if(ikey.equals("CTRL"))
		{
			return "CONTROL";
		}
		if(ikey.equals("BACKSPACE"))
		{
			return "BACK_SPACE";
		}
		return ikey;
	}
	
	public static int getKeyValue(String key) {
		
		int kv = -1;
		Class<KeyEvent>  clazz = KeyEvent.class;
		try {
			key = formatKey(key);
			Field field = clazz.getField("VK_" + key.toUpperCase());
			kv = field.getInt(null);
			return kv;
		} catch (NoSuchFieldException | SecurityException | IllegalArgumentException | IllegalAccessException e) {
			e.printStackTrace();
		}
		return kv;
	}
}

class PressKey{
	
	Robot r = null;
	public PressKey() {
		try {
			r = new Robot();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	// ctrl+ key art + key
	public void Press(int... keys) {
		for(int key:keys) {
			if(key==-1) {
				new Exception("按键不支持：").printStackTrace();
				return;
			}
		}
		try {
			int beginIndex=0;
			for (;beginIndex<keys.length;beginIndex++) {
				r.keyPress(keys[beginIndex]);
			}
			if(keys.length>1) {
				r.delay(10);
			}
			int endIndex = keys.length - 1;
			for (; endIndex >= 0; endIndex--) {
				r.keyRelease(keys[endIndex]);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
