import groovy.text.SimpleTemplateEngine
import groovy.util.slurpersupport.GPathResult
import groovy.util.slurpersupport.NodeChild
import groovy.xml.XmlUtil

import java.util.regex.Pattern
import java.util.zip.ZipEntry
import java.util.zip.ZipFile
import java.util.zip.ZipOutputStream;




class DocReader2 {
	ZipFile zipFileHandle
	SimpleTemplateEngine engine = new SimpleTemplateEngine()
	def commentsMap = [:]
	static main(args) {
		def inputDocFile = "doc2"
		def outputDocFile = "${inputDocFile}_new.docx"
		def binding = ["id" : "000001", "name" : "Vinay Samudre", "fee" : 1000, "payment" : 100, "balance" : 900,
					prodList : [
						[prodId:'PROD-1', prodFee:5000, prodPayment:2000].withDefault { k -> k }
					],
					prodTechSpecs : [
						[prodId:'PROD-1', modelNumber:'M-0001', technicalSpec:'7500RPM Uni Mode'],
						[prodId:'PROD-2', modelNumber:'M-0002', technicalSpec:'7500RPM bi Mode'],
						[prodId:'PROD-3', modelNumber:'M-0003', technicalSpec:'7500RPM multi Mode'],
						[prodId:'PROD-4', modelNumber:'M-0004', technicalSpec:'8500RPM Uni Mode'],
						[prodId:'PROD-5', modelNumber:'M-0005', technicalSpec:'8500RPM bi Mode'],
						[prodId:'PROD-6', modelNumber:'M-0006', technicalSpec:'''8500RPM\n
 multi Mode Recording technology
    DAT 72

Capacity
    72 GB; Maximum, compressed 2:1

Transfer rate
    21.6 GB/hr; Maximum, compressed 2:1

Buffer size
    8 MB; Included

Host interface
    Ultra160 LVD SCSI

Encryption capability
    No

WORM capability
    No

Form factor
    5.25 inch half-height''']
					]
				]
		binding = binding.withDefault { k -> k } //return key as it is if not found in map

		def tab1 = [col1:'COL-1', col2:'COL-2', col3:'COL-3']
		DocReader2 dr = new DocReader2()
		def replacedFileMap = [:] // filename : content
		dr.setZipFileHandle("${inputDocFile}.docx")
		def mainDocFilePath = dr.getMainDocumentFilePath()
		println "MainDocumentFilePath is ${dr.mainDocumentFilePath}"
		def mainDocFileText = dr.readFileFromZip(mainDocFilePath)
		//println "MainDocFileText is ${mainDocFileText}"
		replacedFileMap[mainDocFilePath] = dr.parseAndReplaceXmlDocumentUsingSlurper(mainDocFileText, binding)
		//println "replacedFileMap= $replacedFileMap"
		dr.createWordDoc(outputDocFile, replacedFileMap)
	}

	def createWordDoc(outputDocFile, fileContentMap) {
		ZipOutputStream zout = new ZipOutputStream( new FileOutputStream(outputDocFile) )
		zipFileHandle.entries().each { e ->
			//println "Working on ${e.name}"
			if (fileContentMap[e.name]) {
				println "Creating new Entry ====================================="
				ZipEntry ze = new ZipEntry(e.name)
				zout.putNextEntry(ze)
				zout << fileContentMap[e.name]
			} else {

				ZipEntry ze = new ZipEntry(e.name)
				//println "${ze.name} ${ze.size}"
				zout.putNextEntry(ze)
				zout << zipFileHandle.getInputStream(e)
			}
		}
		zout.close()
	}

	def readFileFromZip(String fileName) {
		String fileText
		zipFileHandle.entries().each { ZipEntry ze ->
			if (fileName == ze.name) {
				println "Checking file : ${ze.name}"
				fileText =  zipFileHandle.getInputStream(ze).text
			}
		}
		return fileText
	}

	def getMainDocumentFilePath() {
		String fileName = "_rels/.rels"
		def fileText = readFileFromZip(fileName)
		GPathResult rootNode = new XmlSlurper(false, true).parseText(fileText)
		print  "*" * 30
		def f = rootNode.Relationship.find { it.@Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" }
		return f.@Target.toString()
	}


	def parseAndReplaceXmlDocumentUsingSlurper(xmlText, binding) {
		GPathResult rootNode = new XmlSlurper(false, true).parseText(xmlText)

		commentsMap = readComments(rootNode)
		println "CommentsMap=$commentsMap"

		println "&" * 50
		println rootNode.getClass()
		rootNode.breadthFirst().each {
			//println "processing ${it.name()}"


			if ( it.name() == 'tbl') {
				processTbl(it, binding)
			} else {
				if ( !isThisIsaTableChild(it, binding) && it.name() == 't' ) {
					println "{$it.name()}   ==>  ${it.text()}"
					def template = engine.createTemplate(it.text())
					def o = template.make(binding)
					println "OLD=${it.text()} NEW=$o"
					it.replaceBody(o.toString())
				}
			}
		}
		return  XmlUtil.serialize(rootNode)
	}



	def processTbl(node, binding) {
		println "Table.........."
		def commentRefIds = readReferencedCommentIds(node)
		def listName = getListName(node, commentRefIds)
		println "listName=$listName"
		def trNode = node.tr[1]
		println trNode.dump()
		trNode.depthFirst().each {
			if ( it.name() =~ /comment/ ) {
				it.replaceNode {}
			}
		}

		def trXml = XmlUtil.serialize(trNode)
		println "Serialized TRNode=" + trXml
		trNode.replaceNode {}
		binding[listName].each { b ->
			def bindingForTR = binding + b

			def clonedTrNode = new XmlSlurper(false, true).parseText(trXml)
			clonedTrNode.breadthFirst().each {
				if ( it.name() == 't' ) {
					println "{$it.name()}   ==>  ${it.text()}"
					def template = engine.createTemplate(it.text())
					def o = template.make(bindingForTR)
					println "TR OLD=${it.text()} NEW=$o"
					it.replaceBody(o.toString())
				}


			}
			trNode.parent().appendNode(clonedTrNode)
		}
	}

	def getListName(node, commentRefIds) {
		def comments = ""
		commentRefIds.each { crid ->
			println "Getting commentId:$crid from $commentsMap"
			def c = commentsMap[crid]
			println "Got $c comment for commentId:$crid from $commentsMap"
			comments += c
			println "comments=$comments"
		}
		println "comments=$comments"
		def commentPropMap = extractCommentPropMap(comments)
		def listName = commentPropMap["listName"]
		return listName
	}

	def extractCommentPropMap(comment) {
		println "extractListName from $comment"
		def keyVals = comment.split(/\s+/)
		def commentPropMap = [:]
		keyVals.each {
			println "Working on comment=$it"


			Pattern p = ~/(\w+)=(\w+)/
			def m = it =~ p
			m.each {
				println it.dump()
				commentPropMap[it[1]] = it[2]
			}
		}
		return commentPropMap
	}

	def readReferencedCommentIds(node) {
		def commentRefIds = node.depthFirst().findAll { it.name() == "commentReference" }
		def commentIds = []

		println "RefIds="
		commentRefIds.each { commentIds << it.'@id'.toString() }
		println commentIds
		return commentIds
	}

	def readComments(node) {
		def referencedCommentIds = readReferencedCommentIds(node)
		def commentsFileText
		def commentsMap = [:]
		if (referencedCommentIds) {
			commentsFileText = readFileFromZip("word/comments.xml")
		}
		println "$commentsFileText"
		GPathResult rootNode = new XmlSlurper(false, true).parseText(commentsFileText)
		rootNode.breadthFirst().each { println "processing ${it.name()}" }



		referencedCommentIds.each { crid ->
			def commentNode = rootNode.comment.findAll { it."@id" == crid }
			commentsMap[crid] = commentNode.text().toString()
		}
		return commentsMap
	}


	def substituteRow(node, binding) {

		node.depthFirst().each {
			if ( it.name() == 't' ) {
				def (listName, columnName) = getListAndColumnName(it.text())
				println "listName=$listname   columnName=$columnName"
				def template = engine.createTemplate(it.text())
				def o = template.make(binding)
				println "OLD=${it.text()} NEW=$o"
				it.replaceBody(o.toString())
			}

		}
	}

	def getListAndColumnName(t) {
		def listAndCol = []
		if ( t ==~ /\$\w+\.\w+/ ){
			listAndCol = t.split(/\./)

		}
		return listAndCol
	}
	def isThisIsaTableChild(NodeChild node, bindings) {
		//println "LIST NODE ${node.getClass()} ${node.name()}   ==>  ${node.text()}"
		def p = node
		def tbl
		while ( p.parent) {
			p = p.parent()

			if ( p.name() == 'tbl') {
				tbl = p
			}
			if ( p.name() == 'document') {
				break
			}
		}
		return tbl
	}

	def setZipFileHandle(String zipFileName) {
		File f = new File(zipFileName)
		println "Path = ${f.absolutePath} ${f.canonicalFile} $f.canonicalPath"
		zipFileHandle = new ZipFile(f)
	}
}


