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
		String path1 = "/Users/Abulkair/Documents/workspace/proj/src/proj/words test.txt";
		String path2 = "/Users/Abulkair/Documents/workspace/proj/src/proj/dict.txt";
		assertEquals(expected, proj.buildChain(path1, path2));
	}
	
	@Test
	public void test2() {
		String path1 = "/Users/Abulkair/Documents/workspace/proj/src/proj/words test 2.txt";
		String path2 = "/Users/Abulkair/Documents/workspace/proj/src/proj/dict.txt";
		assertEquals(expected, proj.buildChain(path1, path2));
	}
	
	@Test
	public void test3() {
		String path1 = "/Users/Abulkair/Documents/workspace/proj/src/proj/words.txt";
		String path2 = "/Users/Abulkair/Documents/workspace/proj/src/proj/dict test.txt";
		assertEquals(expected, proj.buildChain(path1, path2));
	}
	
	@Test
	public void test4() {
		String path1 = "/Users/Abulkair/Documents/workspace/proj/src/proj/words.txt";
		String path2 = "/Users/Abulkair/Documents/workspace/proj/src/proj/dict test 2.txt";
		assertEquals(expected, proj.buildChain(path1, path2));
	}
}
