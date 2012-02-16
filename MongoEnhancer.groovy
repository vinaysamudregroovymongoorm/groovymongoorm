import com.mongodb.Mongo;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.BasicDBObject;
import com.mongodb.DBObject;
import com.mongodb.DBCursor;
import groovy.util.logging.Log4j

@Log4j
class MongoEnhancer {
   def enhanceDomainClass(className) {
      Class c = Class.forName(className)
      c.metaClass.asBasicDBObject = constructBasicDBObject(className)
      c.metaClass.insert          = constructInsertStatement(className)
      c.metaClass.update          = constructUpdateStatement(className)
      c.metaClass.find            = constructFindStatement(className)
      c.metaClass.drop            = constructDropStatement(className)
      c.metaClass.count           = constructCountStatement(className)
      log.info("$className - Enhanced")
   }

   def constructBasicDBObject(className) {
      Class c = Class.forName(className)
      def props = c.metaClass.properties
      props.removeAll { p -> p.name =~ /class|metaClass/ }
      return {
         BasicDBObject doc = new BasicDBObject()
         props.each { 
            def p = delegate["${it.name}"]
            log.trace(" => Working on ${it.name}: $p")
            if ( p != null) {
            switch ( p ) {
               case List:
                 if ( p.size() > 0 && p[0] instanceof BasicMongoEntity ) {
                 log.trace( " is of List<BasicMongoEntity> class" )
                 List l = []
                 delegate["${it.name}"].each {
                    l << it.asBasicDBObject()
                 }
                 doc.put("${it.name}", l) 
                 }
                 break
               case BasicMongoEntity:
                 log.trace( " is of BasicMongoEntity class" )
                 doc.put("${it.name}", p.asBasicDBObject()) 
                 break
               default:
                 log.trace(" is of Standard class")
                    if (  p ==~ /^\*.*/ ||  p ==~ /.*\*$/ ) {
                       def pWithoutAstrisk = ( p =~ /\*/ ).replaceAll("")
                       log.trace("String without * = $pWithoutAstrisk")
                       log.trace("Didnot detect any wildcards")
                       def pPattern = ~/$pWithoutAstrisk/
                       doc.put("${it.name}", pPattern) 
                    } else {
                       log.trace("Didnot detect any wildcards")
                       doc.put("${it.name}", p) 
                       break
                    }
            }
            }
         }
         return doc
      }
   }

   def constructInsertStatement(className) {
      return {
         if ( ! delegate.validate() ) {
            return false
         }
         BasicDBObject doc = delegate.asBasicDBObject()
         DB db = MongoDatabase.mydb
         DBCollection coll = db.getCollection(className)
         log.trace("* Inserting BasicDBObject=$doc")
         coll.insert(doc)
         return true
      }
   }

   def constructUpdateStatement(className) {
      return {
         if ( ! delegate.validate() ) {
            return false
         }
         BasicDBObject doc = delegate.asBasicDBObject()
         DB db = MongoDatabase.mydb
         DBCollection coll = db.getCollection(className)
         log.trace("* Updating BasicDBObject=$doc")
         coll.save(doc) //This uses upsert... so might insert if _id is not set
         return true
      }
   }


   def constructUpsertStatement(className) {
      return {
         if ( ! delegate.validate() ) {
            return false
         }
         BasicDBObject doc = delegate.asBasicDBObject()
         DB db = MongoDatabase.mydb
         DBCollection coll = db.getCollection(className)
         log.trace( "* Upserting BasicDBObject=$doc" )
         coll.save(doc)
         return true
      }
   }

   def constructFindStatement(className) {
      return {
         BasicDBObject doc = delegate.asBasicDBObject()
         DB db = MongoDatabase.mydb
         DBCollection coll = db.getCollection(className)
         log.trace( "Finding $doc in collection $className" )
         def cur = coll.find(doc)
      }
   }

   def constructCountStatement(className) {
      return {
         BasicDBObject doc = delegate.asBasicDBObject()
         DB db = MongoDatabase.mydb
         DBCollection coll = db.getCollection(className)
         log.trace( "Counting $doc in collection $className" )
         return coll.getCount()
      }
   }

   def constructDropStatement(className) {
      return {
         DB db = MongoDatabase.mydb
         DBCollection coll = db.getCollection(className)
         log.trace( "Droping collection $className" )
         def cur = coll.drop()
      }
   }
}
