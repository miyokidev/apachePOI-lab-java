import java.util.*;

public class Language {
	
	public String lang;
	
	public Map<String, String> translations = new HashMap<>();
	
	public Language(String lang) {
		this.lang = lang;
	}
	
	public void addTranslation(String key, String value) {
		translations.put(key, value);
	}
}
