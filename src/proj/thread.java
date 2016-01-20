package proj;

public class thread		//Класс с методом main().
{
	static int i;
//	static ThreadLocal<Integer> i = new ThreadLocal<Integer>();
	
	public static void main(String[] args)
	{
//		System.out.println(i);
		//Создание потока
		Thread myThready = new Thread(new Runnable()
		{
			public void run() //Этот метод будет выполняться в побочном потоке
			{
				Go();
			}
		});
		myThready.setName("1");
		myThready.start();	//Запуск потока

		Thread myThready2 = new Thread(new Runnable()
		{
			public void run() //Этот метод будет выполняться в побочном потоке
			{
				Go();
			}
		});
		myThready2.setName("2");
		myThready2.start();	//Запуск потока
		
//		Go();
		
//		System.out.println(i);
	}
	
	public static void Go() {
		for (int j = 0; j < 10; j++) {
//			System.out.println(Thread.currentThread().getName()+" - "+i);
			try {
				Thread.sleep(0);
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
}