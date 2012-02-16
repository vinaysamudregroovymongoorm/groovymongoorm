import groovy.util.logging.Log4j

@Log4j
class ORMEnhancer {
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
      def tabName = props.remove("tableName")
      return {
         Map doc = [:]
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
         Map doc = delegate.asBasicDBObject()
         def connection = Database.mydb
         log.trace("* Inserting BasicDBObject=$doc")
         connection.insert(doc)
         return true
      }
   }


   def constructFindStatement(className) {
      return {
         BasicDBObject doc = delegate.asBasicDBObject()
         DB db = Database.mydb
         DBCollection coll = db.getCollection(className)
         log.trace( "Finding $doc in collection $className" )
         def cur = coll.find(doc)
      }
   }

   def constructCountStatement(className) {
      return {
         BasicDBObject doc = delegate.asBasicDBObject()
         DB db = Database.mydb
         DBCollection coll = db.getCollection(className)
         log.trace( "Counting $doc in collection $className" )
         return coll.getCount()
      }
   }

   def constructDropStatement(className) {
      return {
         DB db = Database.mydb
         DBCollection coll = db.getCollection(className)
         log.trace( "Droping collection $className" )
         def cur = coll.drop()
      }
   }
}
