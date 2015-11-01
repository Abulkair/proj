package proj;
import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;

public class proj {

	private static String first, second, tmp, tmp2;

	private static ArrayList<String> dict = new ArrayList<String>();
	
	public static void main(String[] args) {
		dict = buildChain("/Users/Abulkair/Documents/workspace/proj/src/proj/words.txt", "/Users/Abulkair/Documents/workspace/proj/src/proj/dict.txt");
		
		if (dict != null){		
			for (String word:dict) {
				System.out.println(word);
			}
		}
	}
	
	public static ArrayList<String> buildChain(String pathWords, String pathDict) {
		ArrayList<String> chain = new ArrayList<String>();
		
		if (!readData(pathWords, pathDict)) return null;
		
		int min=first.length();
		tmp = first;
		
		chain.add(first);
		
		while (getDistance(tmp, second)>1) {
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
				break;
			}
			
			if (tmp==first) {
				System.out.println("Нет слов, отличающиеся на один символ.");
				break;
			}
			
			chain.add(tmp);
		}
		
		chain.add(second);
		
		return chain;
	}
	
	private static boolean readData(String pathWords, String pathDict) {
		try {
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

			//Select words of necessary length
		    while (line != null) {
		        if (line.length()==first.length()) dict.add(line);
		        line = br.readLine();
		    };
		    
		    br.close();
		    
		    if (dict.isEmpty()) {
		    	throw new IOException("Нет слов соответствующей длины в словаре!");
		    }
		    
		    dict.trimToSize();
		} catch (Exception e) {
			System.out.println(e.getMessage());
			return false;
		}
		
		return true;
	}
	
	 public static int getDistance(String s, String t) {
	      if (s == null || t == null) {
	          throw new IllegalArgumentException("Strings must not be null");
	      }

	      int n = s.length(); // length of s
	      int m = t.length(); // length of t

	      if (n == 0) {
	          return m;
	      } else if (m == 0) {
	          return n;
	      }

	      if (n > m) {
	          // swap the input strings to consume less memory
	          String tmp = s;
	          s = t;
	          t = tmp;
	          n = m;
	          m = t.length();
	      }

	      int p[] = new int[n+1]; //'previous' cost array, horizontally
	      int d[] = new int[n+1]; // cost array, horizontally
	      int _d[]; //placeholder to assist in swapping p and d

	      // indexes into strings s and t
	      int i; // iterates through s
	      int j; // iterates through t

	      char t_j; // jth character of t

	      int cost; // cost

	      for (i = 0; i<=n; i++) {
	          p[i] = i;
	      }

	      for (j = 1; j<=m; j++) {
	          t_j = t.charAt(j-1);
	          d[0] = j;

	          for (i=1; i<=n; i++) {
	              cost = s.charAt(i-1)==t_j ? 0 : 1;
	              // minimum of cell to the left+1, to the top+1, diagonally left and up +cost
	              d[i] = Math.min(Math.min(d[i-1]+1, p[i]+1),  p[i-1]+cost);
	          }

	          // copy current distance counts to 'previous row' distance counts
	          _d = p;
	          p = d;
	          d = _d;
	      }

	      // our last action in the above loop was to switch d and p, so p now 
	      // actually has the most recent cost counts
	      return p[n];
	  }
}