import React from "react";
import { motion } from "framer-motion";
import { GraduationCap, Award, Briefcase, MapPin } from "lucide-react";

const stats = [
  { icon: GraduationCap, label: "Degree", value: "BSc Mathematical Sciences", sub: "UNISA – Applied Math & CS", color: "from-violet-500 to-purple-600" },
  { icon: Award, label: "Certificate", value: "Software Development", sub: "Zaio Coding School · Alison", color: "from-cyan-500 to-blue-600" },
  { icon: Award, label: "Certificate", value: "Generative AI", sub: "WeThinkCode", color: "from-pink-500 to-rose-600" },
  { icon: Award, label: "Certificate", value: "Computer Hardware Basics", sub: "Cisco Networking Academy", color: "from-blue-500 to-indigo-600" },
  { icon: Award, label: "Certificate", value: "IT Support", sub: "Alison", color: "from-orange-500 to-amber-600" },
  { icon: Award, label: "Certificate", value: "IT Infrastructure", sub: "Alison", color: "from-yellow-500 to-orange-600" },
  { icon: MapPin, label: "Location", value: "South Africa", sub: "Open to remote work", color: "from-emerald-500 to-teal-600" },
];

export default function AboutSection() {
  return (
    <section id="about" className="py-24 px-6 font-inter relative overflow-hidden">
      <div className="absolute inset-0 bg-mesh opacity-30 pointer-events-none" />

      <div className="max-w-7xl mx-auto relative z-10">
        <motion.div initial={{ opacity: 0, y: 30 }} whileInView={{ opacity: 1, y: 0 }} viewport={{ once: true }}
          transition={{ duration: 0.6 }} className="text-center mb-16">
          <p className="font-mono text-sm tracking-widest uppercase text-violet-400 mb-3">About Me</p>
          <h2 className="text-4xl sm:text-5xl font-bold">
            <span className="text-gradient">Full-Stack Developer</span>
            <span className="text-white"> & Problem Solver</span>
          </h2>
        </motion.div>

        <div className="grid lg:grid-cols-2 gap-16 items-start">
          <motion.div initial={{ opacity: 0, x: -30 }} whileInView={{ opacity: 1, x: 0 }} viewport={{ once: true }}
            transition={{ duration: 0.6 }} className="space-y-5">
            {[
              "Recent Applied Mathematics and Computer Science graduate with strong foundations in mathematics, statistics, and problem-solving.",
              "Completed intensive Software Development bootcamp at Zaio Coding School (Alison-certified) and Generative AI course at WeThinkCode, mastering modern development workflows.",
              "Built multiple full-stack applications using React, Firebase, and modern web technologies with focus on clean code, performance, and user experience.",
              "Eager to apply my technical skills and analytical mindset in a professional environment, contributing to innovative projects and growing as a developer.",
            ].map((text, i) => (
              <p key={i} className="text-white/60 leading-relaxed text-base">{text}</p>
            ))}
          </motion.div>

          <motion.div initial={{ opacity: 0, x: 30 }} whileInView={{ opacity: 1, x: 0 }} viewport={{ once: true }}
            transition={{ duration: 0.6, delay: 0.2 }} className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            {stats.map((stat, i) => (
              <motion.div key={i} whileHover={{ y: -4, scale: 1.02 }} transition={{ type: "spring", stiffness: 300 }}
                className="glass rounded-2xl p-6 border border-white/10 hover:border-white/20 transition-all group cursor-pointer">
                <div className={`w-12 h-12 rounded-xl bg-gradient-to-br ${stat.color} flex items-center justify-center mb-4 shadow-lg group-hover:scale-110 transition-transform`}>
                  <stat.icon className="w-6 h-6 text-white" />
                </div>
                <p className="text-xs text-white/40 uppercase tracking-wider mb-1">{stat.label}</p>
                <p className="text-white font-semibold text-sm">{stat.value}</p>
                <p className="text-white/50 text-xs mt-1">{stat.sub}</p>
              </motion.div>
            ))}
          </motion.div>
        </div>
      </div>
    </section>
  );
}