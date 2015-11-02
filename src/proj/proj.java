package proj;
import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;

public class proj {

	private static String first, second, tmp, tmp2;

	public static ArrayList<String> dict;
	
	public static void main(String[] args) {
		dict = buildChain(args[0], args[1]);
		
		if (dict != null){		
			for (String word:dict) {
				System.out.println(word);
			}
		}
	}
	
	public static ArrayList<String> buildChain(String pathWords, String pathDict) {
		boolean flag = true; 
		ArrayList<String> chain = new ArrayList<String>();
		dict = new ArrayList<String>();
		
		if (!readData(pathWords, pathDict)) return null;
		
		int min=first.length();
		tmp = first;
		
		chain.add(first);
		
		while (getDistance(tmp, second)>1 && flag) {
			tmp2 = tmp;
			for (String word:dict) {
				if (getDistance(tmp, word)==1) {
					if (min > getDistance(word, second)) {
						min = getDistance(word, second);
						tmp = word;
					}
				}
			}
			
			if (tmp==tmp2) {
				System.out.println("Нет слов, отличающиеся на один символ.");
				flag=false;
			}
			
			if (tmp==first) {
				System.out.println("Нет слов, отличающиеся на один символ.");
				flag=false;
			}
			
			if (flag) chain.add(tmp);
		}
		
		chain.add(second);
		
		if (flag) {
			return chain;
		} else {
			return null;
		}
	}
	
	private static boolean readData(String pathWords, String pathDict) {
		try {
			if (pathWords == null || pathDict == null) {
				throw new IOException("Не хватает входных параметров!");
			}
			
			if (Files.notExists(Paths.get(pathWords)) || Files.notExists(Paths.get(pathDict))) {
				throw new IOException("Не существуют один или оба входных файла!");
			}
			
			BufferedReader br = new BufferedReader(new FileReader(pathWords));
		    first = br.readLine();
		    second = br.readLine();
		    br.close();
		    
		    if (first==null) {
		    	throw new IOException("Нет первого слова в первом файле!");
		    }
		    
		    if (second==null) {
		    	throw new IOException("Нет второго слова в первом файле!");
		    }
		    
		    if (first.length()!=second.length()) {
		    	throw new IOException("Слова разной длины!");
		    }
		} catch (Exception e) {
			System.out.println(e.getMessage());
			return false;
		}
		
		try {
			BufferedReader br = new BufferedReader(new FileReader(pathDict));
		    
			String line = br.readLine();

		    while (line != null) {
		        if (line.length()==first.length()) dict.add(line);
		        line = br.readLine();
		    };
		    
		    br.close();
		    
		    if (dict.isEmpty()) {
		    	throw new IOException("Нет слов соответствующей длины в словаре или словарь пуст!");
		    }
		    
		    dict.trimToSize();
		} catch (Exception e) {
			System.out.println(e.getMessage());
			return false;
		}
		
		return true;
	}
	
	//Levenshtein distance
	 public static int getDistance(String s, String t) {
	      if (s == null || t == null) {
	          throw new IllegalArgumentException("Strings must not be null");
	      }

	      int n = s.length();
	      int m = t.length();

	      if (n == 0) {
	          return m;
	      } else if (m == 0) {
	          return n;
	      }

	      if (n > m) {
	          String tmp = s;
	          s = t;
	          t = tmp;
	          n = m;
	          m = t.length();
	      }

	      int p[] = new int[n+1];
	      int d[] = new int[n+1];
	      int _d[];

	      int i;
	      int j;

	      char t_j;

	      int cost;

	      for (i = 0; i<=n; i++) {
	          p[i] = i;
	      }

	      for (j = 1; j<=m; j++) {
	          t_j = t.charAt(j-1);
	          d[0] = j;

	          for (i=1; i<=n; i++) {
	              cost = s.charAt(i-1)==t_j ? 0 : 1;

	              d[i] = Math.min(Math.min(d[i-1]+1, p[i]+1),  p[i-1]+cost);
	          }

	          _d = p;
	          p = d;
	          d = _d;
	      }

	      return p[n];
	  }
}