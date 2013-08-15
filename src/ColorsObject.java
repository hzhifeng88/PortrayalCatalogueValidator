public class ColorsObject{

	private String ColorID;
	private String sRGB;
	private String CMYK;
	private String CIELAB;

	public ColorsObject() {
	}

	public ColorsObject(String ColorID, String sRGB, String CMYK, String CIELAB) {

		this.setColorID(ColorID);
		this.setsRGB(sRGB);
		this.setCMYK(CMYK);
		this.setCIELAB(CIELAB);
	}

	public String getColorID() {
		return ColorID;
	}

	public void setColorID(String colorID) {
		ColorID = colorID;
	}

	public String getsRGB() {
		return sRGB;
	}

	public void setsRGB(String sRGB) {
		this.sRGB = sRGB;
	}

	public String getCMYK() {
		return CMYK;
	}

	public void setCMYK(String cMYK) {
		CMYK = cMYK;
	}

	public String getCIELAB() {
		return CIELAB;
	}

	public void setCIELAB(String cIELAB) {
		CIELAB = cIELAB;
	}
}
