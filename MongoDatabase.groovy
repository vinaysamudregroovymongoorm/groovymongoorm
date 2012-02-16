import com.mongodb.Mongo;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.BasicDBObject;
import com.mongodb.DBObject;
import com.mongodb.DBCursor;

class MongoDatabase {
  def static instances = [:]
  static {
     Mongo m = new Mongo( "localhost" , 27017 );
     DB db = m.getDB( "mydb" );
     instances.mydb = db
  }

  def static getMydb() {
     instances.mydb
  }
}

