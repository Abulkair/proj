package proj;
//
//public class thread2 {
//    private static boolean ready;  
//    private static int number;  
//    private static class ReaderThread extends Thread {   
//        public void run() {  
//            while (!ready)   {
//            	Thread.yield();
//            }
//            System.out.println(number);
//        }  
//    }  
//    
//    public static void main(String[] args) {  
//    	new ReaderThread().start();
//    	float f =0;
//    	int i = 0;
//    	try {
//    	System.out.println(i/i);
//    	} catch (RuntimeException e) {
//    		System.out.println(e.getMessage());
//    	}
//        number = 42;  
//        ready = true;
//    }
//}

abstract public class thread2{ 
    public static void main(String[] args) {
    	Integer i1= 10;
    	int i2= (byte)10;
    	byte b = (int)10;
    	byte i =20;
    	byte j=20;
    	byte k = (byte)(i+j);
    	System.out.println(i2+b);
    	System.out.println(i1==10);
    	System.out.println(5^6&3);
        try { 
            String value = "29.1"; 
            System.out.println((Float.valueOf(value) + 1.0) == 30.1); 
            System.out.println((Double.valueOf(value) + 1.0) == 30.1); 
            System.out.println(Float.valueOf(value)/0); 
            System.out.println(Double.valueOf(value)/0);
        } 
        catch (NumberFormatException ex) { 
            System.out.println("NumberFormatException"); 
        } 
        catch (ArithmeticException ex) { 
            System.out.println("ArithmeticException"); 
        } 
        
        System.out.println(countWaysToProduceGivenAmountOfMoney(-1));
    }
    
    public static int countWaysToProduceGivenAmountOfMoney(int cents) {
    	int cnt = 0;
    	
    	try {
    		if (cents<=0) throw new Exception("Cents should be positive number");
    		
	        int[] dp = new int[cents + 1];
	        dp[0] = 1;
	        for (int x : new int[] {1, 5, 10, 25, 50})
	            for (int i = 0; i + x <= cents; ++i)
	                dp[i + x] += dp[i];
	        
	        cnt = dp[cents];
    	}
    	catch (Exception e) {
    		System.out.println(e.getMessage());
    	}
    	
        return cnt;
    }
} 