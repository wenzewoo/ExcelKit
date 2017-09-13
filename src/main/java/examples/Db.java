/**

 * Copyright (c) 2017, ������ (wuwz@live.com).

 *

 * Licensed under the Apache License, Version 2.0 (the "License");

 * you may not use this file except in compliance with the License.

 * You may obtain a copy of the License at

 *

 *      http://www.apache.org/licenses/LICENSE-2.0

 *

 * Unless required by applicable law or agreed to in writing, software

 * distributed under the License is distributed on an "AS IS" BASIS,

 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.

 * See the License for the specific language governing permissions and

 * limitations under the License.

 */
package examples;

import java.util.ArrayList;
import java.util.List;
import java.util.Random;

public class Db {
	static final List<User> users = new ArrayList<User>();
	
	static {
		// 模拟15条数据
		Random random = new Random();
		
		User user = null;
		for (int i = 0; i < 15; i++) {
			user = new User()
				.setUid(i + 1)
				.setUsername("0000"+(i+1))
				.setPassword("password")
				.setSex(random.nextInt(2) % 2 == 0 ? 1 : 2)
				.setGradeId(random.nextInt(4))
			    .setGendex("male");

			
			users.add(user);
			user = null;
		}
	}

	
	public static List<User> getUsers() {
		return users;
	}
	
	public static void addUser(User user) {
		users.add(user);
	}
}
