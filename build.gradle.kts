/*
 * This file was generated by the Gradle 'init' task.
 *
 * This generated file contains a sample Java project to get you started.
 * For more details take a look at the Java Quickstart chapter in the Gradle
 * User Manual available at https://docs.gradle.org/6.0.1/userguide/tutorial_java_projects.html
 */

plugins {
    // Apply the java plugin to add support for Java
    java

    // Apply the application plugin to add support for building a CLI application.
    application
}

application {
    // Define the main class for the application.
    mainClassName = "org.jid.tests.executeformulaexcel.App"
}

repositories {
    mavenCentral()
    jcenter()
}

dependencies {
    // This dependency is used by the application.
    implementation("org.apache.poi:poi:4.1.1")
    implementation("org.apache.poi:poi-ooxml:4.1.1")

    testImplementation("org.assertj:assertj-core:3.14.0")

    // Use JUnit Jupiter API for testing.
    testImplementation("org.junit.jupiter:junit-jupiter-api:5.5.1")

    // Use JUnit Jupiter Engine for testing.
    testRuntimeOnly("org.junit.jupiter:junit-jupiter-engine:5.5.1")
}

val test by tasks.getting(Test::class) {
    // Use junit platform for unit tests
    useJUnitPlatform()
}
