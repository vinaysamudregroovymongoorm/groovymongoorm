import com.mongodb.Mongo;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.BasicDBObject;
import com.mongodb.DBObject;
import com.mongodb.DBCursor;


def className ="Person"
MongoEnhancer me = new MongoEnhancer()
me.enhanceDomainClass(className)
println new Date()

Person p = new Person()
//p.drop()

setPerson(p, "Vikas8")
Person m = new Person()
setPerson(m, "Vinay")
Person e1 = new Person()
setPerson(e1, "Emp1")
Person e2 = new Person()
setPerson(e2, "Emp2")
p.manager = m
p.employees.add(e1)
p.employees.add(e2)

if (! p.insert() ) {
   println "This object is not VALID"
   println p.errors
}

Person p3 = new Person()
p3.firstName = "Vikas8"
println "COUNT=${p3.count()}"
cur = p3.find()
println cur.class
println "=============== before update start ============="
while(cur.hasNext()) {
  Person m3 = new Person(cur.next().toMap())
  println "...$m3"
}   
println "=============== before update end ${p3.count()} ============="

Person p2 = new Person()
p2.firstName = "Vikas8"
def cur = p2.find()
println cur.class
while(cur.hasNext()) {
  Person m2 = new Person(cur.next().toMap())
  m2.secondName = m2.secondName + " - changed"
  m2.manager.firstName = "Vinay :-}"
  m2.update()
  println "==> updated"
}   

p3 = new Person()
p3.firstName = "Vikas8"
cur = p3.find()
println cur.class
println "=============== after update start ============="
while(cur.hasNext()) {
  Person m3 = new Person(cur.next().toMap())
  println "...$m3"
}   
println "=============== after update end ${p3.count()} ============="

   def setPerson(p, name) {
      p.firstName = name
      p.secondName = name
      p.surName = name
      p.count = 1
      p.hobies = ["a", "b"]
      df = new java.text.SimpleDateFormat("yyyyMMdd")
      p.dob = df.parse("20081029")
      p.comments = [ [ c : "comment1", createdOn : new Date() ],
                     [ c : "comment2", createdOn : new Date() ]
                   ]
      return p
   }
