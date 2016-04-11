javac -sourcepath src -d build src/**/*.java
java -cp bin;lib/* org.testng.TestNG testng.xml
