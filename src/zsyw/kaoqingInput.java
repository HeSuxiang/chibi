package zsyw;

public class kaoqingInput {

	private String name ="无名";
	private float zhourijiabangongshi =0;
	private float qingjiagongshi=0;
	private float qingbingjiagongshi=0;
	private float tiaoxiugongshi=0;
	private int wanjidaka=0;
	

	kaoqingInput(String name){
		this.name = name;
	}
	
	
	
	
	
	
	/**
	 * @return the name
	 */
	public String getName() {
		return name;
	}
	/**
	 * @param name the name to set
	 */
	public void setName(String name) {
		this.name = name;
	}
	/**
	 * @return the zhourijiabangongshi
	 */
	public float getZhourijiabangongshi() {
		return zhourijiabangongshi;
	}
	/**
	 * @param zhourijiabangongshi the zhourijiabangongshi to set
	 */
	public void AddZhourijiabangongshi(float zhourijiabangongshi) {
		this.zhourijiabangongshi += zhourijiabangongshi;
	}
	/**
	 * @return the qingjiagongshi
	 */
	public float getQingjiagongshi() {
		return qingjiagongshi;
	}
	/**
	 * @param qingjiagongshi the qingjiagongshi to set
	 */
	public void AddQingjiagongshi(float qingjiagongshi) {
		this.qingjiagongshi += qingjiagongshi;
	}
	/**
	 * @return the qingbingjiagongshi
	 */
	public float getQingbingjiagongshi() {
		return qingbingjiagongshi;
	}
	/**
	 * @param qingbingjiagongshi the qingbingjiagongshi to set
	 */
	public void AddQingbingjiagongshi(float qingbingjiagongshi) {
		this.qingbingjiagongshi += qingbingjiagongshi;
	}
	
	
	/**
	 * @return the tiaoxiugongshi
	 */
	public float getTiaoxiugongshi() {
		return tiaoxiugongshi;
	}






	/**
	 * @param tiaoxiugongshi the tiaoxiugongshi to set
	 */
	public void AddTiaoxiugongshi(float tiaoxiugongshi) {
		this.tiaoxiugongshi += tiaoxiugongshi;
	}






	/**
	 * @return the wanjidaka
	 */
	public int getWanjidaka() {
		return wanjidaka;
	}
 
	public int AddWanjidaka() {
		wanjidaka +=1;
		return wanjidaka;
	}
}
