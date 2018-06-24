package com.zsf.attendance.entity;

/** 
* @author 张胜锋 
* @date 2018年6月4日 
* @下午9:56:52
* description:员工实体类
*/
public class Staff {

	private String workNum;//工号
	private String name;  //姓名
	private String deptName;  //部门名称
	private float work;
	public String getWorkNum() {
		return workNum;
	}
	public void setWorkNum(String workNum) {
		this.workNum = workNum;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getDeptName() {
		return deptName;
	}
	public void setDeptName(String deptName) {
		this.deptName = deptName;
	}
	public float getWork() {
		return work;
	}
	public void setWork(float work) {
		this.work = work;
	}
	
	
}
