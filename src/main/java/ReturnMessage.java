public class ReturnMessage {
    private boolean success;
    private String ReturnMessageContent;

    public ReturnMessage(boolean success) {
        this.success = success;
    }

    public ReturnMessage() {}

    public String getReturnMessageContent() {
        return ReturnMessageContent;
    }

    public void setReturnMessageContent(String ReturnMessageContent) {
        this.ReturnMessageContent = ReturnMessageContent;
    }

    public boolean isSuccess() {
        return success;
    }

    public void setSuccess(boolean success) {
        this.success = success;
    }
}