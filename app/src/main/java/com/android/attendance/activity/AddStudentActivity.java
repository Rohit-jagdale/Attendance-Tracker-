package com.android.attendance.activity;

import com.android.attendance.bean.StudentBean;
import com.android.attendance.db.DBAdapter;
import com.example.androidattendancesystem.R;

import android.Manifest;
import android.app.Activity;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.graphics.Color;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.support.annotation.NonNull;
import android.text.TextUtils;
import android.util.Log;
import android.view.Menu;
import android.view.View;
import android.view.View.OnClickListener;
import android.widget.AdapterView;
import android.widget.AdapterView.OnItemSelectedListener;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Spinner;
import android.widget.TextView;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class AddStudentActivity extends Activity {

	Button registerButton, createXl;
	EditText textFirstName;
	EditText textLastName;

	EditText textcontact;
	EditText textaddress;
	EditText username;
	EditText password;
	Spinner spinnerbranch,spinneryear;
	String userrole,branch,year;
	ArrayList<StudentBean> StudentList;
	private String[] branchString = new String[] { "cse"};
	private String[] yearString = new String[] {"SE","TE","BE"};

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.addstudent);
		System.setProperty("org.apache.poi.javax.xml.stream.XMLInputFactory", "com.fasterxml.aalto.stax.InputFactoryImpl");
		System.setProperty("org.apache.poi.javax.xml.stream.XMLOutputFactory", "com.fasterxml.aalto.stax.OutputFactoryImpl");
		System.setProperty("org.apache.poi.javax.xml.stream.XMLEventFactory", "com.fasterxml.aalto.stax.EventFactoryImpl");


		spinnerbranch=(Spinner)findViewById(R.id.spinnerdept);
		spinneryear=(Spinner)findViewById(R.id.spinneryear);
		textFirstName=(EditText)findViewById(R.id.editTextFirstName);
		textLastName=(EditText)findViewById(R.id.editTextLastName);
		textcontact=(EditText)findViewById(R.id.editTextPhone);
		textaddress=(EditText)findViewById(R.id.editTextaddr);
		username=(EditText)findViewById(R.id.editTextUserName);
		password=(EditText) findViewById(R.id.editTextPassword);
		registerButton=(Button)findViewById(R.id.RegisterButton);


		spinnerbranch.setOnItemSelectedListener(new OnItemSelectedListener() {
			@Override
			public void onItemSelected(AdapterView<?> arg0, View view,
					int arg2, long arg3) {
				// TODO Auto-generated method stub
				((TextView) arg0.getChildAt(0)).setTextColor(Color.WHITE);
				branch =(String) spinnerbranch.getSelectedItem();

			}

			@Override
			public void onNothingSelected(AdapterView<?> arg0) {
				// TODO Auto-generated method stub
			}
		});

		ArrayAdapter<String> adapter_branch = new ArrayAdapter<String>(this,
				android.R.layout.simple_spinner_item, branchString);
		adapter_branch
		.setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item);
		spinnerbranch.setAdapter(adapter_branch);

		///......................spinner2

		spinneryear.setOnItemSelectedListener(new OnItemSelectedListener() {
			@Override
			public void onItemSelected(AdapterView<?> arg0, View view,
					int arg2, long arg3) {
				// TODO Auto-generated method stub
				((TextView) arg0.getChildAt(0)).setTextColor(Color.WHITE);
				year =(String) spinneryear.getSelectedItem();

			}

			@Override
			public void onNothingSelected(AdapterView<?> arg0) {
				// TODO Auto-generated method stub
			}
		});

		ArrayAdapter<String> adapter_year = new ArrayAdapter<String>(this,
				android.R.layout.simple_spinner_item, yearString);
		adapter_year
		.setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item);
		spinneryear.setAdapter(adapter_year);



		registerButton.setOnClickListener(new OnClickListener() {

			@Override
			public void onClick(View v) {
				// TODO Auto-generated method stub
				//......................................validation
				String first_name = textFirstName.getText().toString();
				String last_name = textLastName.getText().toString();
				String phone_no = textcontact.getText().toString();
				String address = textaddress.getText().toString();
				//String username = username

				if (TextUtils.isEmpty(first_name)) {
					textFirstName.setError("please enter firstname");
				}

				else if (TextUtils.isEmpty(last_name)) {
					textLastName.setError("please enter lastname");
				}
				else if (TextUtils.isEmpty(phone_no)) {
					textcontact.setError("please enter phoneno");
				}

				else if (TextUtils.isEmpty(address)) {
					textaddress.setError("enter address");
				}
				else {

					StudentBean studentBean = new StudentBean();

					studentBean.setStudent_firstname(first_name);
					studentBean.setStudent_lastname(last_name);
					studentBean.setStudent_mobilenumber(phone_no);
					studentBean.setStudent_address(address);
					studentBean.setStudent_department(branch);
					studentBean.setStudent_class(year);

					DBAdapter dbAdapter= new DBAdapter(AddStudentActivity.this);
					dbAdapter.addStudent(studentBean);

					Toast.makeText(getApplicationContext(), "student added successfully", Toast.LENGTH_SHORT).show();

				}
				// TODO: Perform input validation

				// Check if external storage is available and writable
				String state = Environment.getExternalStorageState();
				if (Environment.MEDIA_MOUNTED.equals(state)) {
					// External storage is available and writable, proceed with file creation
					importData(); // This method will handle both student registration and Excel file creation
				} else if (Environment.MEDIA_MOUNTED_READ_ONLY.equals(state)) {
					// External storage is mounted but read-only
					// Inform the user that they need to free up space or ensure the storage is writable
					Toast.makeText(AddStudentActivity.this, "External storage is read-only", Toast.LENGTH_SHORT).show();
				} else {
					// External storage is not mounted
					// Prompt the user to check their device's storage settings or provide alternative options
					Toast.makeText(AddStudentActivity.this, "External storage is not available", Toast.LENGTH_SHORT).show();
				}
				if (Build.VERSION.SDK_INT > Build.VERSION_CODES.M) {
					if (getApplicationContext().checkSelfPermission(Manifest.permission.WRITE_EXTERNAL_STORAGE) == PackageManager.PERMISSION_DENIED) {
						String[] permissions = {Manifest.permission.WRITE_EXTERNAL_STORAGE};
						requestPermissions(permissions, 1);
					} else {
						importData();
					}
				} else {
					importData();
				}
			}
		});

	}
	private void importData() {
		// Create an instance of DBAdapter using the current context
		DBAdapter dbAdapter = new DBAdapter(this);

		// Call getAllStudent() method on the instance of DBAdapter
		StudentList = dbAdapter.getAllStudent(AddStudentActivity.this);


		if (StudentList.size() > 0) {
			createXlFile();
		} else {
			Toast.makeText(this, "List is empty", Toast.LENGTH_SHORT).show();
		}
	}



	private void createXlFile(){
		 File filePath = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS) + "/Demo.xls");

		Workbook wb = new HSSFWorkbook();

		Cell cell = null;

		Sheet sheet = null;
		sheet = wb.createSheet("Demo Excel Sheet");

		//Now column and row
		Row row = sheet.createRow(0);

		cell = row.createCell(0);
		cell.setCellValue("First Name");


		cell = row.createCell(1);
		cell.setCellValue("Last Name");


		cell = row.createCell(2);
		cell.setCellValue("Contact");

		cell =  row.createCell(3);
		cell.setCellValue("Address");

		cell = row.createCell(4);
		cell.setCellValue("department");

		cell =  row.createCell(5);
		cell.setCellValue("year");

		//column width
		sheet.setColumnWidth(0, (20 * 200));
		sheet.setColumnWidth(1, (30 * 200));
		sheet.setColumnWidth(2, (30 * 200));
		sheet.setColumnWidth(3, (30 * 200));
		sheet.setColumnWidth(4, (30 * 200));
		sheet.setColumnWidth(5, (30 * 200));

		for (int i = 0; i < StudentList.size(); i++) {
			Row row1 = sheet.createRow(i + 1);

			cell = row1.createCell(0);
			cell.setCellValue(StudentList.get(i).getStudent_firstname());

			cell = row1.createCell(1);
			cell.setCellValue(StudentList.get(i).getStudent_lastname());
			//  cell.setCellStyle(cellStyle);

			cell = row1.createCell(2);
			cell.setCellValue(StudentList.get(i).getStudent_mobilenumber());

			cell = row1.createCell(3);
			cell.setCellValue(StudentList.get(i).getStudent_address());

			cell = row1.createCell(4);
			cell.setCellValue(StudentList.get(i).getStudent_department());
			//  cell.setCellStyle(cellStyle);

			cell = row1.createCell(5);
			cell.setCellValue(StudentList.get(i).getStudent_class());

			sheet.setColumnWidth(0, (20 * 200));
			sheet.setColumnWidth(1, (30 * 200));
			sheet.setColumnWidth(2, (30 * 200));
			sheet.setColumnWidth(3, (30 * 200));
			sheet.setColumnWidth(4, (30 * 200));
			sheet.setColumnWidth(5, (30 * 200));
		}

		String folderName = "Download";
		String fileName = folderName + System.currentTimeMillis() + ".xls";
		String path = Environment.getExternalStorageDirectory() + File.separator + folderName + File.separator + fileName;

		File file = new File(Environment.getExternalStorageDirectory() + File.separator + folderName);
		if (!file.exists()) {
			file.mkdirs();
		}
		Log.d("FilePath", "File Path: " + path);

		FileOutputStream outputStream = null;

		try {
			outputStream = new FileOutputStream(path);
			wb.write(outputStream);
			Log.d("FileCreation", "Excel File Created Successfully");
			Toast.makeText(getApplicationContext(), "Excel Created in " + path, Toast.LENGTH_SHORT).show();
		} catch (IOException e) {
			e.printStackTrace();
			Log.e("FileCreation", "IOException occurred: " + e.getMessage());
			Toast.makeText(getApplicationContext(), "Failed to create Excel file", Toast.LENGTH_LONG).show();
		} finally {
			try {
				if (outputStream != null) {
					outputStream.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	@Override
	public boolean onCreateOptionsMenu(Menu menu) {
		// Inflate the menu; this adds items to the action bar if it is present.
		getMenuInflater().inflate(R.menu.main, menu);
		return true;
	}
	@Override
	public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
		if (requestCode == 1 && grantResults.length > 0 && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
			importData();
		} else {
			Toast.makeText(getApplicationContext(), "Permission Denied", Toast.LENGTH_SHORT).show();
		}
	}
}
