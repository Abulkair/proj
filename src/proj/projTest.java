package proj;

import static org.junit.Assert.*;

import java.util.ArrayList;

import org.junit.Test;

public class projTest {

    public ArrayList<String> expected = new ArrayList<String>() {{
        add("КОТ");
        add("ТОТ");
        add("ТОН");
    }};
	
	@Test
	public void test1() {
		String path1 = "/Users/Abulkair/Documents/workspace/proj/src/proj/texts/words test.txt";
		String path2 = "/Users/Abulkair/Documents/workspace/proj/src/proj/texts/dict.txt";
		assertEquals(expected, proj.buildChain(path1, path2));
	}
	
	@Test
	public void test2() {
		String path1 = "/Users/Abulkair/Documents/workspace/proj/src/proj/texts/words test 2.txt";
		String path2 = "/Users/Abulkair/Documents/workspace/proj/src/proj/texts/dict.txt";
		assertEquals(expected, proj.buildChain(path1, path2));
	}
	
	@Test
	public void test3() {
		String path1 = "/Users/Abulkair/Documents/workspace/proj/src/proj/texts/words.txt";
		String path2 = "/Users/Abulkair/Documents/workspace/proj/src/proj/texts/dict test.txt";
		assertEquals(expected, proj.buildChain(path1, path2));
	}
	
	@Test
	public void test4() {
		String path1 = "/Users/Abulkair/Documents/workspace/proj/src/proj/texts/words.txt";
		String path2 = "/Users/Abulkair/Documents/workspace/proj/src/proj/texts/dict test 2.txt";
		assertEquals(expected, proj.buildChain(path1, path2));
	}
	
	@Test
	public void test5() {
		String path1 = null;
		String path2 = "";
		assertEquals(expected, proj.buildChain(path1, path2));
	}
	
	@Test
	public void test6() {
		String path1 = "/Users/Abulkair/Documents/workspace/proj/src/proj/texts/words fake.txt";
		String path2 = "/Users/Abulkair/Documents/workspace/proj/src/proj/texts/dict fake.txt";
		assertEquals(expected, proj.buildChain(path1, path2));
	}
}
