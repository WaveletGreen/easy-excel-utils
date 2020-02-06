package excel;

public abstract class BaseTest {
    public String getMethodName() {
        Thread thread = Thread.currentThread();
        String methodName = thread.getStackTrace()[2].getMethodName();
        String fileName = "target/" + methodName + ".xlsx";
        return fileName;
    }
}
