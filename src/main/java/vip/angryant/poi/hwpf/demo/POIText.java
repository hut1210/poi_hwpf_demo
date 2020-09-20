package vip.angryant.poi.hwpf.demo;

public abstract class POIText {
    public abstract String getText();


    public static POIText str(final String string) {
        return new POIText() {
            @Override
            public String getText() {
                return string;
            }
        };
    }
}
