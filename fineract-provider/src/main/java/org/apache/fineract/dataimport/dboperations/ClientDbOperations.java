/**
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements. See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership. The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License. You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied. See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */
package org.apache.fineract.dataimport.dboperations;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

import org.apache.fineract.infrastructure.core.service.RoutingDataSource;
import org.apache.fineract.dataimport.dto.client.Client;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.dao.DataAccessException;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.ResultSetExtractor;
import org.springframework.stereotype.Component;

@Component
public class ClientDbOperations {
	private JdbcTemplate jdbcTemplate;
	
	
	@Autowired
	ClientDbOperations(JdbcTemplate jdbcTemplate, final RoutingDataSource dataSource){
		this.jdbcTemplate=new JdbcTemplate(dataSource);
	}
	
	public List<Client> getClientData(){
		 return jdbcTemplate.query("select * from m_client",new ResultSetExtractor<List<Client>>(){  
			    @Override  
			     public List<Client> extractData(ResultSet rs) throws SQLException,  
			            DataAccessException {  
			      
			        List<Client> list=new ArrayList<>();  
			        while(rs.next()){  
			         Client client=new Client();  
			         client.setRowIndex(rs.getInt(1));
			         client.setOfficeId(String.valueOf(rs.getInt(8)));
			         client.setStaffId(String.valueOf(rs.getInt(10)));
			         client.setFirstname(rs.getString(11));
			         client.setMiddlename(rs.getString(12));
			         client.setLastname(rs.getString(13));
			         client.setFullname(rs.getString(14));
			         client.setExternalId(rs.getString(3));
			         client.setActivationDate(String.valueOf(rs.getDate(6)));
			         list.add(client);  
			        }  
			        return list;  
			        }  
			    }); 
	}
	
}
