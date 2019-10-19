package com.anuj.utilization.client.net;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.net.InetAddress;
import java.net.InetSocketAddress;
import java.net.Socket;
import java.net.SocketAddress;
import java.net.SocketTimeoutException;
import java.net.UnknownHostException;

public class OpenSocket {

	static Socket sock;
	 static String server = "10.108.0.33";
	 static int port = 5550;
	 static String filename = "/foo/bar/application1.log";
	 static String command = "tail -50 " + filename + "\n";
	public static void main(String[] args) {
	    openSocket();
	    try 
	    {
	      // write to socket
	     /* BufferedWriter wr = new BufferedWriter(new OutputStreamWriter(sock.getOutputStream()));
	      wr.write(command);
	      wr.flush();*/
	      
	      // read from socket
	      BufferedReader rd = new BufferedReader(new InputStreamReader(sock.getInputStream()));
	      String str;
	      while ((str = rd.readLine()) != null)
	      {
	        System.out.println(str);
	      }
	      rd.close();
	    } 
	    catch (IOException e) 
	    {
	      System.err.println(e);
	    }
	  }


	private static void openSocket()
	  {
	    // open a socket and connect with a timeout limit
	    try
	    {
	      InetAddress addr = InetAddress.getByName(server);
	      SocketAddress sockaddr = new InetSocketAddress(addr, port);
	      sock = new Socket();
	  
	      // this method will block for the defined number of milliseconds
	      int timeout = 2000;
	      sock.connect(sockaddr, timeout);
	    } 
	    catch (UnknownHostException e) 
	    {
	      e.printStackTrace();
	    }
	    catch (SocketTimeoutException e) 
	    {
	      e.printStackTrace();
	    }
	    catch (IOException e) 
	    {
	      e.printStackTrace();
	    }
	  }
}
