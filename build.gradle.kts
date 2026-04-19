import org.jetbrains.compose.desktop.application.dsl.TargetFormat
import org.jetbrains.kotlin.gradle.dsl.JvmTarget
import org.jetbrains.kotlin.gradle.tasks.KotlinCompile

plugins {
    kotlin("jvm")
    id("org.jetbrains.compose")
    id("org.jetbrains.kotlin.plugin.compose")
    id("com.github.johnrengelman.shadow") version "8.1.1"
}

group = "org.example"
version = "1.0-SNAPSHOT"

repositories {
    mavenCentral()
    maven("https://maven.pkg.jetbrains.space/public/p/compose/dev")
    google()
}

dependencies {
    val composeVersion = "1.7.0"
    // Compose Desktop platform artifact (includes core UI)
    implementation(compose.desktop.currentOs)

    // Common Compose libraries (use plugin-provided coordinates when available)
    // These aliases resolve with the org.jetbrains.compose plugin and provide
    // the commonly-used UI libraries for desktop apps.
    implementation("org.jetbrains.compose.material3:material3-desktop:$composeVersion")
    implementation(compose.ui)
    implementation(compose.foundation)
    implementation(compose.runtime)
    implementation(compose.animation)
    implementation(compose.preview)
    implementation(compose.uiTooling)

    // Apache POI for Excel handling
    implementation("org.apache.poi:poi:5.4.0")
    implementation("org.apache.poi:poi-ooxml:5.4.0")

    // Logging
    implementation("org.slf4j:slf4j-simple:2.0.13")

    testImplementation(kotlin("test"))
}

kotlin {
    jvmToolchain(21)
}

tasks.withType<KotlinCompile>().configureEach {
    compilerOptions {
        jvmTarget.set(JvmTarget.JVM_21)
    }
}

compose.desktop {
    application {
        mainClass = "MainKt"

        nativeDistributions {
            targetFormats(TargetFormat.Dmg, TargetFormat.Msi, TargetFormat.Deb, TargetFormat.Exe)
            packageName = "ExcelAggregator"
            packageVersion = "1.0.0"
        }
    }
}


// Make the uber jar part of the build lifecycle
tasks.named("build") {
    dependsOn("shadowJar")
}

tasks.test {
    useJUnitPlatform()
}
